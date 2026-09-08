import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  downloadSubtitleGenerationModel,
  generateJapaneseSubtitles,
  resolveSubtitleGenerationModel,
} from '../../src/core/services/subtitle-generation.js';
import {
  resolveSubtitleGenerationConfig,
  type SubtitleGenerationProgress,
} from '../../src/shared/subtitle-generation.js';
import {
  readLauncherMainConfigObject,
  resolveLauncherMainConfigPath,
} from '../config/shared-config-reader.js';
import { sendMpvCommandWithResponse } from '../mpv.js';
import { resolvePathMaybe } from '../util.js';
import type { LauncherCommandContext } from './context.js';

type GenerationCommandContext = Pick<LauncherCommandContext, 'args' | 'mpvSocketPath'> & {
  processAdapter: Pick<LauncherCommandContext['processAdapter'], 'writeStdout' | 'setExitCode'>;
};

interface GenerationCommandDeps {
  readConfig: typeof readLauncherMainConfigObject;
  configPath: typeof resolveLauncherMainConfigPath;
  resolveModel: typeof resolveSubtitleGenerationModel;
  downloadModel: typeof downloadSubtitleGenerationModel;
  generate: typeof generateJapaneseSubtitles;
  mpvCommand: typeof sendMpvCommandWithResponse;
  onInterrupt: (handler: () => void) => () => void;
}

const defaultDeps: GenerationCommandDeps = {
  readConfig: readLauncherMainConfigObject,
  configPath: resolveLauncherMainConfigPath,
  resolveModel: resolveSubtitleGenerationModel,
  downloadModel: downloadSubtitleGenerationModel,
  generate: generateJapaneseSubtitles,
  mpvCommand: sendMpvCommandWithResponse,
  onInterrupt: (handler) => {
    process.on('SIGINT', handler);
    return () => process.off('SIGINT', handler);
  },
};

function localMediaPath(value: string, workingDirectory = process.cwd()): string {
  if (value.startsWith('file://')) return fileURLToPath(value);
  if (/^[a-z][a-z\d+.-]*:\/\//i.test(value)) {
    throw new Error('Japanese subtitle generation requires a local media file.');
  }
  return path.resolve(workingDirectory, resolvePathMaybe(value));
}

async function readMpvMedia(socketPath: string, command: GenerationCommandDeps['mpvCommand']) {
  const media = await command(socketPath, ['get_property', 'path'], 1000);
  if (typeof media !== 'string' || !media.trim()) return null;
  let workingDirectory: string | undefined;
  if (!path.isAbsolute(media) && !media.startsWith('file://')) {
    const directory = await command(socketPath, ['get_property', 'working-directory'], 1000);
    if (typeof directory !== 'string') return null;
    workingDirectory = directory;
  }
  return localMediaPath(media, workingDirectory);
}

async function readMpvAudioStream(
  socketPath: string,
  command: GenerationCommandDeps['mpvCommand'],
) {
  const tracks = await command(socketPath, ['get_property', 'track-list'], 1000);
  for (const track of Array.isArray(tracks) ? tracks : []) {
    if (
      typeof track === 'object' &&
      track !== null &&
      'type' in track &&
      track.type === 'audio' &&
      'selected' in track &&
      track.selected === true
    ) {
      if ('external' in track && track.external === true) {
        throw new Error(
          'The selected mpv audio track is external. Pass its local file to generate-subs.',
        );
      }
      if (
        'ff-index' in track &&
        typeof track['ff-index'] === 'number' &&
        Number.isSafeInteger(track['ff-index']) &&
        track['ff-index'] >= 0
      ) {
        return track['ff-index'];
      }
    }
  }
  throw new Error(
    'Select an audio track in mpv, or pass --audio-stream with its absolute stream index.',
  );
}

function sameFile(left: string, right: string): boolean {
  try {
    return fs.realpathSync(left) === fs.realpathSync(right);
  } catch {
    return path.resolve(left) === path.resolve(right);
  }
}

/** Keep progress readable in terminals and redirected logs, even for large model downloads. */
export function createGenerationProgressReporter(write: (text: string) => void, now = Date.now) {
  let previousStage: SubtitleGenerationProgress['stage'] | undefined;
  let previousTime = -Infinity;
  let previousLine = '';
  return (progress: SubtitleGenerationProgress): void => {
    const percent =
      typeof progress.percent === 'number' && Number.isFinite(progress.percent)
        ? Math.floor(Math.max(0, Math.min(100, progress.percent)))
        : undefined;
    const line = `[${progress.stage}] ${percent === undefined ? '' : `${percent}% `}${progress.message}\n`;
    const timestamp = now();
    if (
      line === previousLine ||
      (progress.stage === previousStage && timestamp - previousTime < 1000 && percent !== 100)
    )
      return;
    write(line);
    previousStage = progress.stage;
    previousTime = timestamp;
    previousLine = line;
  };
}

export async function runGenerateSubtitlesCommand(
  context: GenerationCommandContext,
  overrides: Partial<GenerationCommandDeps> = {},
): Promise<boolean> {
  const options = context.args.generateSubtitles;
  if (!options) return false;
  const deps = { ...defaultDeps, ...overrides };
  const write = (text: string) => context.processAdapter.writeStdout(text);
  const controller = new AbortController();
  const removeInterrupt = deps.onInterrupt(() => controller.abort());
  try {
    const config = resolveSubtitleGenerationConfig(deps.readConfig()?.subtitleGeneration);
    if (options.managedModel) {
      config.managedModel = options.managedModel;
      config.modelPath = '';
    }
    if (options.modelPath !== undefined)
      config.modelPath = path.resolve(resolvePathMaybe(options.modelPath));
    const modelDirectory = path.join(path.dirname(deps.configPath()), 'models', 'whisper');
    const currentMedia = await readMpvMedia(context.mpvSocketPath, deps.mpvCommand).catch(
      () => null,
    );
    const mediaPath = options.mediaPath ? localMediaPath(options.mediaPath) : currentMedia;
    if (!mediaPath)
      throw new Error('Pass a local video file or open one in mpv before running generate-subs.');
    const audioStreamIndex =
      options.audioStreamIndex ??
      (!options.mediaPath
        ? await readMpvAudioStream(context.mpvSocketPath, deps.mpvCommand)
        : undefined);
    const onProgress = createGenerationProgressReporter(write);
    const model = await deps.resolveModel(config, modelDirectory);
    if (model.kind === 'invalid') throw new Error(model.message);
    if (model.kind === 'missing') {
      if (!options.downloadModel) {
        throw new Error(
          'No Whisper model found. Run again with --download-model, or set subtitleGeneration.modelPath / --model-path.',
        );
      }
      await deps.downloadModel({ config, modelDirectory, onProgress, signal: controller.signal });
    }
    const outputPath = await deps.generate({
      config,
      modelDirectory,
      mediaPath,
      audioStreamIndex,
      outputPath: options.outputPath
        ? path.resolve(resolvePathMaybe(options.outputPath))
        : undefined,
      onProgress,
      signal: controller.signal,
    });
    write(`Saved Japanese subtitles: ${outputPath}\n`);
    controller.signal.throwIfAborted();
    const playingMedia = await readMpvMedia(context.mpvSocketPath, deps.mpvCommand).catch(
      () => null,
    );
    controller.signal.throwIfAborted();
    if (playingMedia && sameFile(playingMedia, mediaPath)) {
      try {
        await deps.mpvCommand(context.mpvSocketPath, [
          'sub-add',
          outputPath,
          'select',
          'Japanese (generated)',
          'ja',
        ]);
        await deps.mpvCommand(context.mpvSocketPath, ['set_property', 'sub-delay', 0]);
        write('Loaded Japanese subtitles into mpv.\n');
      } catch (error) {
        write(
          `Subtitles are saved, but mpv could not load them: ${error instanceof Error ? error.message : String(error)}\n`,
        );
        context.processAdapter.setExitCode(1);
      }
    }
    return true;
  } catch (error) {
    if (!controller.signal.aborted) throw error;
    write('Subtitle generation cancelled.\n');
    context.processAdapter.setExitCode(130);
    return true;
  } finally {
    removeInterrupt();
  }
}
