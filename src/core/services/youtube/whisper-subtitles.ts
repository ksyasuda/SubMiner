import { mkdtemp, readdir, rm } from 'node:fs/promises';
import { rmSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type {
  SubtitleGenerationConfig,
  SubtitleGenerationProgress,
} from '../../../shared/subtitle-generation';
import { generateJapaneseSubtitles, requireSubtitleGenerationSetup } from '../subtitle-generation';
import { runSubtitleGenerationProcess } from '../subtitle-generation-process';
import { getYoutubeYtDlpCommand, YTDLP_SINGLE_VIDEO_ARG } from './ytdlp-command';

export const YOUTUBE_WHISPER_AUDIO_DIR_PREFIX = 'subminer-youtube-audio-';

/**
 * Whisper resamples everything to 16 kHz mono, so bitrate beyond plain speech clarity only costs
 * download time. Take the smallest audio-only stream of at least 48 kbps, falling back to the
 * smallest one. Sorting on `lang` first keeps the original track ahead of auto-dubbed ones.
 */
const AUDIO_FORMAT = 'ba[abr>=48]/ba';
const AUDIO_SORT = 'lang,+abr';

export function buildYoutubeWhisperAudioArgs(url: string, outputTemplate: string): string[] {
  return [
    YTDLP_SINGLE_VIDEO_ARG,
    '--no-warnings',
    '--newline',
    // Override user yt-dlp config that would change the audio or add requests: SponsorBlock
    // removal shifts every subtitle after a cut, and subtitle embedding hits the timedtext API.
    '--no-sponsorblock',
    '--no-write-subs',
    '--no-write-auto-subs',
    '--no-embed-subs',
    '--no-write-thumbnail',
    '--no-embed-thumbnail',
    '--no-embed-metadata',
    '--no-embed-chapters',
    '-f',
    AUDIO_FORMAT,
    '-S',
    AUDIO_SORT,
    '-o',
    outputTemplate,
    url,
  ];
}

export function parseYtDlpDownloadPercent(line: string): number | null {
  const match = /^\[download\]\s+(\d+(?:\.\d+)?)%/.exec(line.trim());
  return match ? Math.min(100, Number(match[1])) : null;
}

async function findDownloadedAudio(directory: string): Promise<string> {
  const names = await readdir(directory);
  const audio = names.find((name) => name.startsWith('audio.') && !/\.(part|ytdl)$/.test(name));
  if (!audio) throw new Error('yt-dlp finished without saving the audio.');
  return path.join(directory, audio);
}

export type YoutubeWhisperSubtitleServiceDeps = {
  getConfig: () => SubtitleGenerationConfig;
  getModelDirectory: () => string;
  tempRoot?: string;
  getYtDlpCommand?: () => string;
  requireSetup?: typeof requireSubtitleGenerationSetup;
  runProcess?: typeof runSubtitleGenerationProcess;
  generate?: typeof generateJapaneseSubtitles;
};

/**
 * Generates Japanese subtitles for a YouTube video by transcribing its audio with Whisper.
 * The audio lives in a private temp directory that is removed as soon as generation ends,
 * whether it succeeds, fails, or is aborted.
 */
export function createYoutubeWhisperSubtitleService(deps: YoutubeWhisperSubtitleServiceDeps) {
  const tempRoot = deps.tempRoot ?? os.tmpdir();
  const activeDirs = new Set<string>();

  // Audio directories only outlive a job when the app crashes mid-generation.
  const removeStaleAudioDirs = async (): Promise<void> => {
    const names = await readdir(tempRoot).catch(() => []);
    await Promise.all(
      names
        .filter((name) => name.startsWith(YOUTUBE_WHISPER_AUDIO_DIR_PREFIX))
        .map((name) => path.join(tempRoot, name))
        .filter((dir) => !activeDirs.has(dir))
        .map((dir) => rm(dir, { recursive: true, force: true }).catch(() => undefined)),
    );
  };

  async function generate(input: {
    url: string;
    outputPath: string;
    /** Overrides the configured settings, e.g. the modal's per-session model choice. */
    config?: SubtitleGenerationConfig;
    signal?: AbortSignal;
    onProgress?: (progress: SubtitleGenerationProgress) => void;
  }): Promise<string> {
    input.signal?.throwIfAborted();
    const config = input.config ?? deps.getConfig();
    // Fail before spending time on the download when Whisper cannot run.
    await (deps.requireSetup ?? requireSubtitleGenerationSetup)(config, deps.getModelDirectory());
    await removeStaleAudioDirs();
    const directory = await mkdtemp(path.join(tempRoot, YOUTUBE_WHISPER_AUDIO_DIR_PREFIX));
    activeDirs.add(directory);
    try {
      input.onProgress?.({ stage: 'download', percent: 0, message: 'Downloading audio...' });
      await (deps.runProcess ?? runSubtitleGenerationProcess)({
        command: (deps.getYtDlpCommand ?? getYoutubeYtDlpCommand)(),
        args: buildYoutubeWhisperAudioArgs(input.url, path.join(directory, 'audio.%(ext)s')),
        signal: input.signal,
        missingMessage: 'yt-dlp was not found. Install it or set SUBMINER_YTDLP_BIN.',
        onLine: (line) => {
          const percent = parseYtDlpDownloadPercent(line);
          if (percent !== null)
            input.onProgress?.({ stage: 'download', percent, message: 'Downloading audio...' });
        },
      });
      return await (deps.generate ?? generateJapaneseSubtitles)({
        config,
        modelDirectory: deps.getModelDirectory(),
        mediaPath: await findDownloadedAudio(directory),
        outputPath: input.outputPath,
        // Keep the extracted WAV inside the managed directory so every cleanup path removes it.
        workDirectory: directory,
        onProgress: input.onProgress,
        signal: input.signal,
      });
    } finally {
      await rm(directory, { recursive: true, force: true }).catch(() => undefined);
      activeDirs.delete(directory);
    }
  }

  /** Synchronous so it still runs while the app is quitting. */
  function removeActiveAudio(): void {
    for (const directory of activeDirs) {
      try {
        rmSync(directory, { recursive: true, force: true });
      } catch {
        // Leftovers are swept before the next download.
      }
    }
    activeDirs.clear();
  }

  return { generate, removeActiveAudio };
}
