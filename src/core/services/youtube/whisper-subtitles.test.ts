import assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import test from 'node:test';
import { DEFAULT_SUBTITLE_GENERATION_CONFIG } from '../../../shared/subtitle-generation';
import {
  buildYoutubeWhisperAudioArgs,
  createYoutubeWhisperSubtitleService,
  parseYtDlpDownloadPercent,
  YOUTUBE_WHISPER_AUDIO_DIR_PREFIX,
} from './whisper-subtitles';

const URL = 'https://www.youtube.com/watch?v=abcdefghijk';

function makeTempRoot(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-youtube-whisper-test-'));
}

function audioDirs(root: string): string[] {
  return fs.readdirSync(root).filter((name) => name.startsWith(YOUTUBE_WHISPER_AUDIO_DIR_PREFIX));
}

/** Stands in for yt-dlp: writes the file its `-o` template names and reports progress. */
function fakeYtDlp(calls: string[][]) {
  return async (input: { args: string[]; onLine?: (line: string) => void }) => {
    calls.push(input.args);
    const template = input.args[input.args.indexOf('-o') + 1]!;
    fs.writeFileSync(template.replace('%(ext)s', 'm4a'), 'audio');
    input.onLine?.('[download]  50.0% of 3.89MiB at 1.00MiB/s ETA 00:02');
    return '';
  };
}

function createService(root: string, overrides: Record<string, unknown> = {}) {
  return createYoutubeWhisperSubtitleService({
    getConfig: () => DEFAULT_SUBTITLE_GENERATION_CONFIG,
    getModelDirectory: () => '/models',
    tempRoot: root,
    getYtDlpCommand: () => 'yt-dlp',
    requireSetup: async () => ({
      modelPath: '/models/ggml-small.bin',
      tools: { ffmpeg: 'ffmpeg', ffprobe: 'ffprobe', whisper: 'whisper-cli', vad: null },
    }),
    ...overrides,
  });
}

test('audio download picks a small original-language stream and ignores user embeds', () => {
  const args = buildYoutubeWhisperAudioArgs(URL, '/tmp/audio.%(ext)s');
  assert.deepEqual(args.slice(args.indexOf('-f'), args.indexOf('-f') + 4), [
    '-f',
    'ba[abr>=48]/ba',
    '-S',
    'lang,+abr',
  ]);
  for (const flag of ['--no-playlist', '--no-sponsorblock', '--no-embed-subs']) {
    assert.ok(args.includes(flag), flag);
  }
  assert.equal(args.at(-1), URL);
});

test('parseYtDlpDownloadPercent reads yt-dlp progress lines only', () => {
  assert.equal(parseYtDlpDownloadPercent('[download]  42.5% of 3.89MiB'), 42.5);
  assert.equal(parseYtDlpDownloadPercent('[download] Destination: audio.m4a'), null);
});

test('generate transcribes the downloaded audio and deletes it afterwards', async () => {
  const root = makeTempRoot();
  const calls: string[][] = [];
  const progress: string[] = [];
  let transcribedPath = '';
  const service = createService(root, {
    runProcess: fakeYtDlp(calls),
    generate: async (input: { mediaPath: string; outputPath: string; workDirectory?: string }) => {
      transcribedPath = input.mediaPath;
      assert.equal(fs.readFileSync(input.mediaPath, 'utf8'), 'audio');
      // The extracted WAV must land where cleanup already looks.
      assert.equal(input.workDirectory, path.dirname(input.mediaPath));
      return input.outputPath;
    },
  });

  const result = await service.generate({
    url: URL,
    outputPath: '/subs/youtube-whisper.ja.srt',
    onProgress: (update) => progress.push(`${update.stage}:${update.percent}`),
  });

  assert.equal(result, '/subs/youtube-whisper.ja.srt');
  assert.equal(calls.length, 1);
  assert.equal(path.basename(transcribedPath), 'audio.m4a');
  assert.deepEqual(progress, ['download:0', 'download:50']);
  assert.deepEqual(audioDirs(root), []);
});

test('generate deletes the audio when transcription fails', async () => {
  const root = makeTempRoot();
  const service = createService(root, {
    runProcess: fakeYtDlp([]),
    generate: async () => {
      throw new Error('whisper crashed');
    },
  });

  await assert.rejects(
    service.generate({ url: URL, outputPath: '/subs/x.srt' }),
    /whisper crashed/,
  );
  assert.deepEqual(audioDirs(root), []);
});

test('generate skips the download when Whisper is not set up', async () => {
  const root = makeTempRoot();
  const calls: string[][] = [];
  const service = createService(root, {
    runProcess: fakeYtDlp(calls),
    requireSetup: async () => {
      throw new Error('No Whisper model found.');
    },
  });

  await assert.rejects(
    service.generate({ url: URL, outputPath: '/subs/x.srt' }),
    /No Whisper model/,
  );
  assert.equal(calls.length, 0);
  assert.deepEqual(audioDirs(root), []);
});

test('leftover audio from a crashed run is removed before the next download', async () => {
  const root = makeTempRoot();
  fs.mkdirSync(path.join(root, `${YOUTUBE_WHISPER_AUDIO_DIR_PREFIX}stale`));
  let dirsDuringDownload: string[] = [];
  const service = createService(root, {
    runProcess: async (input: { args: string[] }) => {
      dirsDuringDownload = audioDirs(root);
      return fakeYtDlp([])(input);
    },
    generate: async (input: { outputPath: string }) => input.outputPath,
  });

  await service.generate({ url: URL, outputPath: '/subs/x.srt' });
  assert.equal(dirsDuringDownload.length, 1);
  assert.notEqual(dirsDuringDownload[0], `${YOUTUBE_WHISPER_AUDIO_DIR_PREFIX}stale`);
});

test('removeActiveAudio deletes audio for a job that is still running', async () => {
  const root = makeTempRoot();
  let release: (() => void) | null = null;
  const service = createService(root, {
    runProcess: fakeYtDlp([]),
    generate: async (input: { outputPath: string }) => {
      await new Promise<void>((resolve) => {
        release = resolve;
      });
      return input.outputPath;
    },
  });

  const job = service.generate({ url: URL, outputPath: '/subs/x.srt' });
  while (!release) await new Promise((resolve) => setImmediate(resolve));
  assert.equal(audioDirs(root).length, 1);
  service.removeActiveAudio();
  assert.deepEqual(audioDirs(root), []);
  (release as () => void)();
  await job;
});
