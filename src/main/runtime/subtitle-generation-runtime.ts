import path from 'node:path';
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
import { extractYoutubeVideoId, toYoutubeWatchUrl } from './youtube-playback';

interface GenerationMpvClient {
  connected: boolean;
  requestProperty: (name: string) => Promise<unknown>;
  request: (command: unknown[]) => Promise<{ error?: string }>;
}

export interface SubtitleGenerationRuntimeDeps {
  getConfig: () => SubtitleGenerationConfig;
  getModelDirectory: () => string;
  getMpvClient: () => GenerationMpvClient | null;
  onProgress: (progress: SubtitleGenerationProgress) => void;
  generate?: typeof generateJapaneseSubtitles;
  download?: typeof downloadSubtitleGenerationModel;
  resolveModel?: typeof resolveSubtitleGenerationModel;
  resolveTools?: typeof resolveSubtitleGenerationTools;
  detectAcceleration?: typeof detectSubtitleGenerationAcceleration;
  downloadVad?: typeof downloadSubtitleGenerationVadModel;
  resolveVadModel?: typeof resolveSubtitleGenerationVadModel;
  /**
   * Transcribes a YouTube video's audio into a temporary subtitle file and returns its path.
   * Without it the modal only generates for local media.
   */
  generateYoutube?: (input: {
    url: string;
    config: SubtitleGenerationConfig;
    signal: AbortSignal;
    onProgress: (progress: SubtitleGenerationProgress) => void;
  }) => Promise<string>;
  /**
   * Maps a playing stream URL back to its YouTube page URL. Windows YouTube playback opens a
   * resolved direct stream in mpv, so mpv's path alone does not identify the video.
   */
  resolveYoutubeSourceUrl?: (mediaPath: string) => string | null;
}

async function currentLocalMedia(client: GenerationMpvClient | null): Promise<string | null> {
  if (!client?.connected) return null;
  const media = await client.requestProperty('path');
  if (typeof media !== 'string' || !media || /^[a-z][a-z\d+.-]*:\/\//i.test(media)) return null;
  if (path.isAbsolute(media)) return path.normalize(media);
  const directory = await client.requestProperty('working-directory');
  return typeof directory === 'string' ? path.resolve(directory, media) : null;
}

async function loadGeneratedSubtitles(client: GenerationMpvClient, outputPath: string) {
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
}

function selectedAudioIndex(tracks: unknown): number {
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
    if ('external' in track && track.external === true)
      throw new Error('Select an audio track inside the local video before generating subtitles.');
    if (
      'ff-index' in track &&
      typeof track['ff-index'] === 'number' &&
      Number.isInteger(track['ff-index']) &&
      track['ff-index'] >= 0
    )
      return track['ff-index'];
    throw new Error('The selected audio track has no FFmpeg stream index.');
  }
  throw new Error('Select an audio track in mpv before generating subtitles.');
}

export function createSubtitleGenerationRuntime(deps: SubtitleGenerationRuntimeDeps) {
  let controller: AbortController | null = null;
  let progress: SubtitleGenerationProgress | null = null;
  let lastResult: SubtitleGenerationResult | null = null;
  let selectedModel: SubtitleGenerationModelId | null = null;
  let vadEnabled: boolean | null = null;
  // Shown as the modal's media while a YouTube job runs; YouTube is never local media.
  let youtubeJobUrl: string | null = null;
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
  // The YouTube page behind a resolved stream URL, or the media path itself.
  function youtubeSource(mediaPath: string | null | undefined): string | null | undefined {
    return (mediaPath && deps.resolveYoutubeSourceUrl?.(mediaPath)) || mediaPath;
  }
  async function currentYoutubeMedia(client: GenerationMpvClient | null): Promise<string | null> {
    if (!client?.connected) return null;
    const media = await client.requestProperty('path');
    return typeof media === 'string' ? toYoutubeWatchUrl(youtubeSource(media)) : null;
  }
  const report = (update: SubtitleGenerationProgress) => {
    progress = update;
    deps.onProgress(update);
  };

  async function run(
    operation: (signal: AbortSignal) => Promise<SubtitleGenerationResult>,
  ): Promise<SubtitleGenerationResult> {
    if (controller)
      return { ok: false, message: 'A subtitle generation or model download is already running.' };
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
    const client = deps.getMpvClient();
    const mediaPath =
      youtubeJobUrl ??
      (await currentLocalMedia(client).catch(() => null)) ??
      (deps.generateYoutube ? await currentYoutubeMedia(client).catch(() => null) : null);
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

  // YouTube subtitles only exist in mpv, so the video pauses while they generate and resumes
  // afterwards. Closing the modal leaves the job running; unpausing is up to the user.
  async function startYoutube(
    client: GenerationMpvClient,
    url: string,
    config: SubtitleGenerationConfig,
    signal: AbortSignal,
    generate: NonNullable<SubtitleGenerationRuntimeDeps['generateYoutube']>,
  ): Promise<SubtitleGenerationResult> {
    const resume = (await client.requestProperty('pause')) === false;
    if (resume) await client.request(['set_property', 'pause', true]);
    youtubeJobUrl = url;
    const stillPlaying = async () =>
      deps.getMpvClient() === client &&
      (await currentYoutubeMedia(client).catch(() => null)) === url;
    try {
      const outputPath = await generate({ url, config, signal, onProgress: report });
      if (signal.aborted || !(await stillPlaying()))
        return { ok: false, message: 'Playback changed, so the subtitles were not loaded.' };
      await loadGeneratedSubtitles(client, outputPath);
      return { ok: true, outputPath, message: 'Japanese subtitles generated and loaded.' };
    } finally {
      youtubeJobUrl = null;
      if (resume && (await stillPlaying()))
        await client.request(['set_property', 'pause', false]).catch(() => undefined);
    }
  }

  return {
    getStatus,
    /** Stops a YouTube job once its video is no longer playing, which also deletes its audio. */
    handleMediaPathChange(mediaPath: string | null | undefined): void {
      if (!youtubeJobUrl) return;
      if (extractYoutubeVideoId(youtubeSource(mediaPath)) !== extractYoutubeVideoId(youtubeJobUrl))
        controller?.abort();
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
    /**
     * Runs YouTube Whisper generation as the active operation, so the modal shows its progress
     * and its Cancel button stops it. Aborting `signal` stops it too. Resolves to the written
     * subtitle file, or null when the job was cancelled.
     */
    async generateYoutubeSubtitles(input: {
      url: string;
      signal: AbortSignal;
      generate: (
        signal: AbortSignal,
        onProgress: (progress: SubtitleGenerationProgress) => void,
      ) => Promise<string>;
    }): Promise<string | null> {
      let cancelled = false;
      let outputPath: string | null = null;
      const result = await run(async (signal) => {
        const stop = () => controller?.abort();
        input.signal.addEventListener('abort', stop, { once: true });
        if (input.signal.aborted) stop();
        youtubeJobUrl = input.url;
        try {
          outputPath = await input.generate(signal, report);
          return { ok: true, outputPath, message: 'Japanese subtitles generated.' };
        } finally {
          cancelled = signal.aborted;
          input.signal.removeEventListener('abort', stop);
          youtubeJobUrl = null;
        }
      });
      if (result.ok) return outputPath;
      if (cancelled) return null;
      throw new Error(result.message);
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
        const youtubeUrl = deps.generateYoutube ? await currentYoutubeMedia(client) : null;
        if (client && youtubeUrl && deps.generateYoutube)
          return await startYoutube(client, youtubeUrl, config, signal, deps.generateYoutube);
        const mediaPath = await currentLocalMedia(client);
        if (!client || !mediaPath)
          throw new Error('Open a local video or audio file, or a YouTube video, in mpv first.');
        const tracks = await client.requestProperty('track-list');
        const audioStreamIndex = selectedAudioIndex(tracks);
        const references = await readSubtitleGenerationReferences(tracks, (name) =>
          client.requestProperty(name),
        );
        if ((await currentLocalMedia(client)) !== mediaPath)
          throw new Error('The current media changed. Start generation again.');
        signal.throwIfAborted();
        const outputPath = await (deps.generate ?? generateJapaneseSubtitles)({
          config,
          modelDirectory: deps.getModelDirectory(),
          mediaPath,
          audioStreamIndex,
          references,
          onProgress: report,
          signal,
        });
        // Saving succeeds even if playback changes or disconnects during the job.
        try {
          const playingMedia = await currentLocalMedia(client);
          if (!signal.aborted && deps.getMpvClient() === client && playingMedia === mediaPath) {
            await loadGeneratedSubtitles(client, outputPath);
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
