import path from 'node:path';
import { createSubtitleGenerationSession } from '../../core/services/subtitle-generation-session';
import type {
  SubtitleGenerationAlternative,
  SubtitleGenerationRemoteSource,
} from '../../core/services/subtitle-generation-source';
import { resolveMpvHttpHeaders } from '../../core/services/mpv-http-headers';
import { readSubtitleGenerationReferences } from '../../core/services/subtitle-generation-reference';
import { detectSubtitleGenerationAcceleration } from '../../core/services/subtitle-generation-acceleration';
import { SUBTITLE_GENERATION_VAD_MODEL } from '../../shared/subtitle-generation-vad-model';
import {
  downloadSubtitleGenerationVadModel,
  resolveSubtitleGenerationVadModel,
} from '../../core/services/subtitle-generation-vad-model';
import type { SubtitleGenerationModelId } from '../../shared/subtitle-generation-model-catalog';
import {
  downloadSubtitleGenerationModel,
  generateJapaneseSubtitles,
  resolveSubtitleGenerationModel,
  resolveSubtitleGenerationTools,
} from '../../core/services/subtitle-generation';
import type {
  SubtitleGenerationConfig,
  SubtitleGenerationProgress,
} from '../../shared/subtitle-generation';
import type {
  SubtitleGenerationResult,
  SubtitleGenerationStatus,
} from '../../shared/subtitle-generation-ipc';

interface GenerationMpvClient {
  connected: boolean;
  requestProperty: (name: string) => Promise<unknown>;
  request: (command: unknown[]) => Promise<{ error?: string }>;
}

export interface SubtitleGenerationRuntimeDeps {
  session?: ReturnType<typeof createSubtitleGenerationSession>;
  getConfig: () => SubtitleGenerationConfig;
  getModelDirectory: () => string;
  getCacheDirectory: () => string;
  getMpvClient: () => GenerationMpvClient | null;
  getAlternativeSources?: (mediaPath: string) => readonly SubtitleGenerationAlternative[];
  onProgress: (progress: SubtitleGenerationProgress) => void;
  generate?: typeof generateJapaneseSubtitles;
  download?: typeof downloadSubtitleGenerationModel;
  resolveModel?: typeof resolveSubtitleGenerationModel;
  resolveTools?: typeof resolveSubtitleGenerationTools;
  detectAcceleration?: typeof detectSubtitleGenerationAcceleration;
  downloadVad?: typeof downloadSubtitleGenerationVadModel;
  resolveVadModel?: typeof resolveSubtitleGenerationVadModel;
}

async function currentMedia(client: GenerationMpvClient | null): Promise<string | null> {
  if (!client?.connected) return null;
  const media = await client.requestProperty('path');
  if (typeof media !== 'string' || !media) return null;
  if (/^https?:\/\//i.test(media)) return media;
  if (/^[a-z][a-z\d+.-]*:\/\//i.test(media)) return null;
  if (path.isAbsolute(media)) return path.normalize(media);
  const directory = await client.requestProperty('working-directory');
  return typeof directory === 'string' ? path.resolve(directory, media) : null;
}

function selectedAudio(
  tracks: unknown,
): { kind: 'internal'; index: number } | { kind: 'external'; url: string; index?: number } {
  if (!Array.isArray(tracks)) throw new Error('Unable to inspect the selected audio track.');
  for (const track of tracks) {
    if (
      !track ||
      typeof track !== 'object' ||
      !('type' in track) ||
      track.type !== 'audio' ||
      !('selected' in track) ||
      track.selected !== true
    )
      continue;
    if ('external' in track && track.external === true) {
      if (
        !('external-filename' in track) ||
        typeof track['external-filename'] !== 'string' ||
        !/^https?:\/\//i.test(track['external-filename'])
      )
        throw new Error(
          'Select an audio track inside the video or an HTTP audio stream before generating subtitles.',
        );
      return {
        kind: 'external',
        url: track['external-filename'],
        ...('ff-index' in track &&
        typeof track['ff-index'] === 'number' &&
        Number.isInteger(track['ff-index']) &&
        track['ff-index'] >= 0
          ? { index: track['ff-index'] }
          : {}),
      };
    }
    if (
      'ff-index' in track &&
      typeof track['ff-index'] === 'number' &&
      Number.isInteger(track['ff-index']) &&
      track['ff-index'] >= 0
    )
      return { kind: 'internal', index: track['ff-index'] };
    throw new Error('The selected audio track has no FFmpeg stream index.');
  }
  throw new Error('Select an audio track in mpv before generating subtitles.');
}

export function createSubtitleGenerationRuntime(deps: SubtitleGenerationRuntimeDeps) {
  const session = deps.session ?? createSubtitleGenerationSession(deps.getCacheDirectory());
  let closed = false;
  let finished: Promise<void> = Promise.resolve();
  let controller: AbortController | null = null;
  let progress: SubtitleGenerationProgress | null = null;
  let lastResult: SubtitleGenerationResult | null = null;
  let selectedModel: SubtitleGenerationModelId | null = null;
  let vadEnabled: boolean | null = null;
  let accelerationCheck:
    | {
        path: string;
        expires: number;
        result: ReturnType<typeof detectSubtitleGenerationAcceleration>;
      }
    | undefined;
  function getConfig(): SubtitleGenerationConfig {
    const config = deps.getConfig();
    return {
      ...config,
      managedModel: selectedModel ?? config.managedModel,
      vadModelPath:
        vadEnabled === null
          ? config.vadModelPath
          : vadEnabled
            ? config.vadModelPath.trim() ||
              path.resolve(deps.getModelDirectory(), SUBTITLE_GENERATION_VAD_MODEL.filename)
            : '',
    };
  }
  const report = (update: SubtitleGenerationProgress) => {
    progress = update;
    deps.onProgress(update);
  };

  async function run(
    operation: (signal: AbortSignal) => Promise<SubtitleGenerationResult>,
  ): Promise<SubtitleGenerationResult> {
    if (closed) return { ok: false, message: 'Subtitle generation session has closed.' };
    if (controller)
      return { ok: false, message: 'A subtitle generation or model download is already running.' };
    let resolveFinished = () => {};
    finished = new Promise<void>((resolve) => {
      resolveFinished = resolve;
    });
    const active = new AbortController();
    controller = active;
    progress = null;
    lastResult = null;
    try {
      lastResult = await operation(active.signal);
    } catch (error) {
      lastResult = {
        ok: false,
        message: active.signal.aborted
          ? 'Cancelled.'
          : error instanceof Error
            ? error.message
            : String(error),
      };
    } finally {
      controller = null;
      resolveFinished();
    }
    return lastResult;
  }

  async function getStatus(): Promise<SubtitleGenerationStatus> {
    const config = getConfig();
    const tools = await (deps.resolveTools ?? resolveSubtitleGenerationTools)(config);
    const whisperPath = tools.whisper.kind === 'found' ? tools.whisper.path : '';
    if (
      !accelerationCheck ||
      accelerationCheck.path !== whisperPath ||
      (!controller && Date.now() >= accelerationCheck.expires)
    ) {
      accelerationCheck = {
        path: whisperPath,
        expires: Date.now() + 30_000,
        result: (deps.detectAcceleration ?? detectSubtitleGenerationAcceleration)(tools.whisper),
      };
    }
    const acceleration = await accelerationCheck.result;
    const model = await (deps.resolveModel ?? resolveSubtitleGenerationModel)(
      config,
      deps.getModelDirectory(),
    );
    const mediaPath = await currentMedia(deps.getMpvClient()).catch(() => null);
    return {
      model,
      vad: {
        enabled: Boolean(config.vadModelPath.trim()),
        model: await (deps.resolveVadModel ?? resolveSubtitleGenerationVadModel)(
          deps.getConfig(),
          deps.getModelDirectory(),
        ),
      },
      // Session toggles decide whether the speech detector executable is required.
      tools,
      acceleration,
      managedModel: config.managedModel,
      externalModelPath: config.modelPath.trim() || null,
      mediaPath,
      running: controller !== null,
      progress,
      lastResult,
    };
  }

  return {
    getStatus,
    initialize: () => session.directory(),
    async dispose(): Promise<void> {
      closed = true;
      controller?.abort();
      await finished;
      await session.dispose();
    },
    async setVadEnabled(enabled: boolean): Promise<SubtitleGenerationStatus> {
      if (controller)
        throw new Error('Wait for the current operation before changing speech detection.');
      vadEnabled = enabled;
      lastResult = null;
      progress = null;
      return getStatus();
    },
    downloadVad(): Promise<SubtitleGenerationResult> {
      return run(async (signal) => {
        await (deps.downloadVad ?? downloadSubtitleGenerationVadModel)({
          config: deps.getConfig(),
          modelDirectory: deps.getModelDirectory(),
          onProgress: report,
          signal,
        });
        return { ok: true, message: 'Speech detection model is ready.' };
      });
    },
    async selectModel(model: SubtitleGenerationModelId): Promise<SubtitleGenerationStatus> {
      if (controller) throw new Error('Wait for the current operation before changing models.');
      if (deps.getConfig().modelPath.trim())
        throw new Error('Clear Model Path in Settings before choosing a managed model.');
      selectedModel = model;
      lastResult = null;
      progress = null;
      return getStatus();
    },
    cancel(): void {
      controller?.abort();
    },
    download(): Promise<SubtitleGenerationResult> {
      return run(async (signal) => {
        await (deps.download ?? downloadSubtitleGenerationModel)({
          config: getConfig(),
          modelDirectory: deps.getModelDirectory(),
          onProgress: report,
          signal,
        });
        return { ok: true, message: 'Model downloaded. Ready to generate Japanese subtitles.' };
      });
    },
    start(): Promise<SubtitleGenerationResult> {
      return run(async (signal) => {
        const config = getConfig();
        if (config.vadModelPath.trim()) {
          const vad = await (deps.resolveVadModel ?? resolveSubtitleGenerationVadModel)(
            deps.getConfig(),
            deps.getModelDirectory(),
          );
          if (vad.kind === 'missing')
            throw new Error(
              'Download the optional speech detection model or turn off Focus on spoken dialogue.',
            );
          if (vad.kind === 'invalid') throw new Error(vad.message);
        }
        const client = deps.getMpvClient();
        const mediaPath = await currentMedia(client);
        if (!client || !mediaPath)
          throw new Error('Open a local media file or HTTP stream in mpv first.');
        const tracks = await client.requestProperty('track-list');
        const audio = selectedAudio(tracks);
        const audioStreamIndex = audio.kind === 'internal' ? audio.index : undefined;
        if (audio.kind === 'external' && !/^https?:\/\//i.test(mediaPath))
          throw new Error('External audio generation requires an HTTP episode stream.');
        const remote: SubtitleGenerationRemoteSource | undefined = /^https?:\/\//i.test(mediaPath)
          ? {
              httpHeaders: await resolveMpvHttpHeaders(client),
              cacheDirectory: deps.getCacheDirectory(),
              sessionDirectory: await session.directory(),
              ...(deps.getAlternativeSources
                ? { alternatives: deps.getAlternativeSources(mediaPath) }
                : {}),
            }
          : undefined;
        if (audio.kind === 'external' && remote) {
          const delay = await client.requestProperty('audio-delay');
          if (typeof delay !== 'number' || !Number.isFinite(delay))
            throw new Error('Unable to inspect the selected audio delay.');
          const matched = remote.alternatives?.find((candidate) => candidate.url === audio.url);
          remote.selectedAudio = {
            url: audio.url,
            audioStreamIndex: audio.index,
            delaySeconds: delay,
            httpHeaders: matched?.httpHeaders ?? remote.httpHeaders,
          };
        }
        const references = await readSubtitleGenerationReferences(tracks, (name) =>
          client.requestProperty(name),
        );
        if ((await currentMedia(client)) !== mediaPath)
          throw new Error('The current media changed. Start generation again.');
        signal.throwIfAborted();
        const outputPath = await (deps.generate ?? generateJapaneseSubtitles)({
          config,
          modelDirectory: deps.getModelDirectory(),
          mediaPath,
          audioStreamIndex,
          remote,
          references,
          onProgress: report,
          signal,
        });
        // Saving succeeds even if playback changes or disconnects during the job.
        try {
          const playingMedia = await currentMedia(client);
          if (!signal.aborted && deps.getMpvClient() === client && playingMedia === mediaPath) {
            const loaded = await client.request([
              'sub-add',
              outputPath,
              'select',
              'Generated Japanese',
              'ja',
            ]);
            if (loaded.error && loaded.error !== 'success') throw new Error(loaded.error);
            const delay = await client.request(['set_property', 'sub-delay', 0]);
            if (delay.error && delay.error !== 'success') throw new Error(delay.error);
            return {
              ok: true,
              outputPath,
              message: `Japanese subtitles saved and loaded: ${outputPath}`,
            };
          }
          return {
            ok: true,
            outputPath,
            message: `Subtitles saved: ${outputPath}. ${signal.aborted ? 'Cancelled after saving; the file was not loaded.' : 'Playback changed, so the file was not loaded.'}`,
          };
        } catch (error) {
          return {
            ok: true,
            outputPath,
            message: `Subtitles saved: ${outputPath}. Could not finish loading into mpv: ${error instanceof Error ? error.message : String(error)}`,
          };
        }
      });
    },
  };
}
