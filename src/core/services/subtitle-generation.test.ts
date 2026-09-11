import assert from 'node:assert/strict';
import { chmod, mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  DEFAULT_SUBTITLE_GENERATION_CONFIG,
  resolveSubtitleGenerationConfig,
  type SubtitleGenerationProgress,
} from '../../shared/subtitle-generation';
import {
  downloadSubtitleGenerationModel,
  generateJapaneseSubtitles,
  resolveSubtitleGenerationModel,
} from './subtitle-generation';
import { runSubtitleGenerationProcess } from './subtitle-generation-process';

async function fixture(run: (directory: string) => Promise<void>) {
  const directory = await mkdtemp(path.join(tmpdir(), 'subtitle-generation-test-'));
  try {
    await run(directory);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
}

async function executable(directory: string, name: string, body: string) {
  const file = path.join(directory, name);
  await writeFile(file, `#!${process.execPath}\n${body}`, { mode: 0o755 });
  return file;
}

function modelHeader(vocabularySize = 51865): Buffer {
  const header = Buffer.alloc(8);
  header.writeUInt32LE(0x67676d6c, 0);
  header.writeInt32LE(vocabularySize, 4);
  return header;
}

async function generationFixture(directory: string) {
  const modelPath = path.join(directory, 'external.bin');
  const mediaPath = path.join(directory, 'episode.mkv');
  const callsPath = path.join(directory, 'calls.jsonl');
  await writeFile(modelPath, modelHeader());
  await writeFile(mediaPath, 'local media');
  const record = `require('node:fs').appendFileSync(${JSON.stringify(callsPath)}, JSON.stringify(process.argv.slice(2)) + '\\n');`;
  const ffprobePath = await executable(
    directory,
    'ffprobe',
    `${record}\nprocess.stdout.write(JSON.stringify({streams: [{index:1,codec_type:'audio',start_time:'10',tags:{language:'eng'}},{index:3,codec_type:'audio',start_time:'12.5',duration:'20',tags:{language:'jpn'}}],format:{start_time:'10',duration:'25'}}));`,
  );
  const ffmpegPath = await executable(
    directory,
    'ffmpeg',
    `${record}\nrequire('node:fs').writeFileSync(process.argv.at(-1), 'wav'); process.stdout.write('out_time_'); setTimeout(() => process.stdout.write('us=10000000\\nprogress=end\\n'), 10);`,
  );
  const whisperPath = await executable(
    directory,
    'whisper-cli',
    `${record}\nconst args=process.argv.slice(2); require('node:fs').writeFileSync(args[args.indexOf('-of')+1]+'.srt', '1\\n00:00:01,000 --> 00:00:02,000\\nこんにちは\\n'); process.stderr.write('whisper_print_progress_callback: progress = '); setTimeout(() => process.stderr.write('55%\\n'), 10);`,
  );
  return {
    config: {
      ...DEFAULT_SUBTITLE_GENERATION_CONFIG,
      modelPath,
      ffprobePath,
      ffmpegPath,
      whisperPath,
    },
    mediaPath,
    modelDirectory: path.join(directory, 'models'),
    callsPath,
  };
}

test('config parser accepts supported models and rejects unsafe threads and wrong field types', () => {
  const warnings: string[] = [];
  const result = resolveSubtitleGenerationConfig(
    { modelPath: '/tmp/whisper.bin', threads: 0, whisperPath: 42, managedModel: 'large-v3-turbo' },
    (key) => warnings.push(key),
  );
  assert.equal(result.modelPath, '/tmp/whisper.bin');
  assert.equal(result.managedModel, 'large-v3-turbo');
  assert.equal(result.threads, DEFAULT_SUBTITLE_GENERATION_CONFIG.threads);
  assert.deepEqual(warnings, ['whisperPath', 'threads']);
});

test('external model path wins and invalid external models never fall back to download', () =>
  fixture(async (directory) => {
    const input = await generationFixture(directory);
    assert.deepEqual(await resolveSubtitleGenerationModel(input.config, input.modelDirectory), {
      kind: 'external',
      path: input.config.modelPath,
    });
    const invalid = { ...input.config, modelPath: path.join(directory, 'missing.bin') };
    assert.equal(
      (await resolveSubtitleGenerationModel(invalid, input.modelDirectory)).kind,
      'invalid',
    );
    await assert.rejects(
      downloadSubtitleGenerationModel({ ...input, config: invalid }),
      /Cannot read model/,
    );
    assert.equal(
      (
        await resolveSubtitleGenerationModel(
          { ...input.config, modelPath: '' },
          input.modelDirectory,
        )
      ).kind,
      'missing',
    );
  }));

test('English-only and incompatible external models are rejected before transcription', () =>
  fixture(async (directory) => {
    const input = await generationFixture(directory);
    await writeFile(input.config.modelPath, modelHeader(51864));
    const englishOnly = await resolveSubtitleGenerationModel(input.config, input.modelDirectory);
    assert.equal(englishOnly.kind, 'invalid');
    assert.ok('message' in englishOnly);
    assert.match(englishOnly.message, /English-only/);
    await assert.rejects(generateJapaneseSubtitles(input), /requires a multilingual model/);
    await writeFile(input.config.modelPath, 'not a GGML model');
    await assert.rejects(generateJapaneseSubtitles(input), /Unsupported model format/);
    await writeFile(input.config.modelPath, modelHeader().subarray(0, 4));
    await assert.rejects(generateJapaneseSubtitles(input), /Unsupported model format/);
    await assert.rejects(readFile(input.callsPath), /ENOENT/);
  }));

test('generation picks Japanese audio, restores timeline offsets, reports split progress, and preserves existing output', () =>
  fixture(async (directory) => {
    const input = await generationFixture(directory);
    const existing = path.join(directory, 'episode.ja.generated.srt');
    await writeFile(existing, 'user subtitles');
    const progress: SubtitleGenerationProgress[] = [];
    const result = await generateJapaneseSubtitles({
      ...input,
      onProgress: (event) => progress.push(event),
    });
    assert.equal(result, path.join(directory, 'episode.ja.generated.1.srt'));
    assert.equal(await readFile(existing, 'utf8'), 'user subtitles');
    assert.match(await readFile(result, 'utf8'), /00:00:03,500 --> 00:00:04,500\nこんにちは/);
    const calls = (await readFile(input.callsPath, 'utf8'))
      .trim()
      .split('\n')
      .map((line): unknown => JSON.parse(line));
    assert.ok(Array.isArray(calls[1]));
    assert.ok(calls[1].includes('0:3'));
    assert.ok(Array.isArray(calls[2]));
    assert.ok(calls[2].includes('ja'));
    assert.ok(calls[2].includes('-osrt'));
    assert.ok(progress.some((event) => event.stage === 'extract' && event.percent === 50));
    assert.ok(progress.some((event) => event.stage === 'transcribe' && event.percent === 55));
    assert.deepEqual(
      (await readdir(directory)).filter((file) => file.startsWith('.subminer-')),
      [],
    );
  }));

test('explicit audio stream and output path are respected without overwriting existing files', () =>
  fixture(async (directory) => {
    const input = await generationFixture(directory);
    const outputPath = path.join(directory, 'chosen.srt');
    const result = await generateJapaneseSubtitles({ ...input, audioStreamIndex: 1, outputPath });
    assert.equal(result, outputPath);
    assert.match(await readFile(result, 'utf8'), /00:00:01,000 --> 00:00:02,000/);
    await assert.rejects(generateJapaneseSubtitles({ ...input, outputPath }), /already exists/);
    await assert.rejects(
      generateJapaneseSubtitles({ ...input, audioStreamIndex: 99 }),
      /stream 99 was not found/,
    );
  }));

test('dialogue generation keeps separate speech passages on the media timeline', () =>
  fixture(async (directory) => {
    const input = await generationFixture(directory);
    const vadModelPath = path.join(directory, 'vad.bin');
    await writeFile(vadModelPath, 'speech detector model');
    const vadPath = await executable(
      directory,
      'vad',
      "process.stdout.write('Detected 2 speech segments:\\nSpeech segment 0: start = 1000.00, end = 1100.00\\nSpeech segment 1: start = 10000.00, end = 10100.00\\n');",
    );
    const whisperPath = await executable(
      directory,
      'dialogue-whisper',
      "const args=process.argv.slice(2); for(let i=0;i<args.length;i++) if(args[i]==='-of') require('node:fs').writeFileSync(args[i+1]+'.srt', '1\\n00:00:00,000 --> 00:01:39,000\\nはい\\n');",
    );
    const output = await generateJapaneseSubtitles({
      ...input,
      config: { ...input.config, vadModelPath, vadPath, whisperPath },
    });
    assert.equal(
      await readFile(output, 'utf8'),
      '1\n00:00:12,500 --> 00:00:13,500\nはい\n\n2\n00:01:42,500 --> 00:01:43,500\nはい\n',
    );
  }));

test('dialogue generation uses quiet pauses and stitches overlapping chunks on the media timeline', () =>
  fixture(async (directory) => {
    const input = await generationFixture(directory);
    const vadModelPath = path.join(directory, 'vad.bin');
    await writeFile(vadModelPath, 'speech detector model');
    const vadPath = await executable(
      directory,
      'vad',
      `const assert = require('node:assert/strict');
const args = process.argv.slice(2);
assert.equal(args[args.indexOf('--vad-min-speech-duration-ms') + 1], '100');
assert.equal(args[args.indexOf('-vp') + 1], '350');
process.stdout.write('Detected 1 speech segments:\\nSpeech segment 0: start = 1000.00, end = 4500.00\\n');`,
    );
    const ffmpegPath = await executable(
      directory,
      'pause-ffmpeg',
      `const args = process.argv.slice(2);
if (args.includes('-af')) {
  process.stderr.write('[silencedetect] silence_end: 28.1 | silence_duration: 0.2\\n');
} else {
  require('node:fs').writeFileSync(args.at(-1), 'wav');
}`,
    );
    const whisperPath = await executable(
      directory,
      'overlap-whisper',
      `const args = process.argv.slice(2);
for (let i = 0; i < args.length; i++) if (args[i] === '-of') {
  const time = args[i + 1].endsWith('speech-0')
    ? '00:00:17,800 --> 00:00:18,250'
    : '00:00:00,100 --> 00:00:00,650';
  require('node:fs').writeFileSync(args[i + 1] + '.srt', '1\\n' + time + '\\nはい\\n');
}`,
    );
    const output = await generateJapaneseSubtitles({
      ...input,
      config: { ...input.config, vadModelPath, vadPath, ffmpegPath, whisperPath },
    });
    assert.equal(await readFile(output, 'utf8'), '1\n00:00:30,300 --> 00:00:30,900\nはい\n');
  }));

test('no detected speech stops generation without transcribing the full audio', () =>
  fixture(async (directory) => {
    const input = await generationFixture(directory);
    const vadModelPath = path.join(directory, 'vad.bin');
    await writeFile(vadModelPath, 'speech detector model');
    const vadPath = await executable(
      directory,
      'vad',
      "process.stdout.write('Detected 0 speech segments:\\n');",
    );
    await assert.rejects(
      generateJapaneseSubtitles({ ...input, config: { ...input.config, vadModelPath, vadPath } }),
      /No spoken dialogue detected/,
    );
    assert.equal((await readFile(input.callsPath, 'utf8')).trim().split('\n').length, 2);
    assert.deepEqual(
      (await readdir(directory)).filter((file) => file.endsWith('.srt')),
      [],
    );
  }));

test('empty executable paths find tools on PATH and explicit overrides take precedence', () =>
  fixture(async (directory) => {
    const input = await generationFixture(directory);
    const previousPath = process.env.PATH;
    process.env.PATH = directory;
    try {
      const config = {
        ...DEFAULT_SUBTITLE_GENERATION_CONFIG,
        modelPath: input.config.modelPath,
      };
      const result = await generateJapaneseSubtitles({ ...input, config });
      assert.match(await readFile(result, 'utf8'), /こんにちは/);
      await assert.rejects(
        generateJapaneseSubtitles({
          ...input,
          config: { ...config, ffprobePath: path.join(directory, 'missing-override') },
        }),
        /missing-override \(subtitleGeneration\.ffprobePath\) is not an executable file/,
      );
    } finally {
      if (previousPath === undefined) delete process.env.PATH;
      else process.env.PATH = previousPath;
    }
  }));

test('generation rejects remote media, missing models, and missing tools before starting a subprocess', () =>
  fixture(async (directory) => {
    const input = await generationFixture(directory);
    await assert.rejects(
      generateJapaneseSubtitles({ ...input, mediaPath: 'https://example.com/movie.mkv' }),
      /local media file/,
    );
    await assert.rejects(
      generateJapaneseSubtitles({ ...input, config: { ...input.config, modelPath: '' } }),
      /No Whisper model found/,
    );
    await assert.rejects(
      generateJapaneseSubtitles({
        ...input,
        config: {
          ...input.config,
          vadModelPath: path.join(directory, 'vad.bin'),
          vadPath: path.join(directory, 'missing-detector'),
        },
      }),
      /missing-detector \(subtitleGeneration\.vadPath\) is not an executable file/,
    );
    await assert.rejects(readFile(input.callsPath), /ENOENT/);
  }));

test(
  'generation rejects an unwritable destination before extracting audio',
  {
    skip: process.platform === 'win32' || process.getuid?.() === 0,
  },
  () =>
    fixture(async (directory) => {
      const input = await generationFixture(directory);
      const readOnly = path.join(directory, 'read-only');
      await mkdir(readOnly, { mode: 0o555 });
      try {
        await assert.rejects(
          generateJapaneseSubtitles({ ...input, outputPath: path.join(readOnly, 'out.srt') }),
          /read-only is not writable/,
        );
        await assert.rejects(readFile(input.callsPath), /ENOENT/);
      } finally {
        await chmod(readOnly, 0o755);
      }
    }),
);

test('process cancellation terminates work and bounds diagnostic output', () =>
  fixture(async (directory) => {
    const slow = await executable(
      directory,
      'slow',
      "process.stdout.write('ready\\n');setInterval(()=>{},1000);",
    );
    const controller = new AbortController();
    await assert.rejects(
      runSubtitleGenerationProcess({
        command: slow,
        args: [],
        signal: controller.signal,
        onLine: () => controller.abort(),
      }),
      /cancelled/,
    );
    const failed = await executable(
      directory,
      'failed',
      "process.stderr.write('x'.repeat(100000));process.exitCode=7;",
    );
    await assert.rejects(
      runSubtitleGenerationProcess({ command: failed, args: [] }),
      (error: unknown) =>
        error instanceof Error &&
        error.message.length < 66000 &&
        error.message.includes('status 7'),
    );
  }));

test('download uses the pinned model revision and removes files that fail integrity', () =>
  fixture(async (directory) => {
    const originalFetch = globalThis.fetch;
    globalThis.fetch = Object.assign(async (request: string | URL | Request) => {
      assert.equal(
        request,
        'https://huggingface.co/ggerganov/whisper.cpp/resolve/5359861c739e955e79d9a303bcbc70fb988958b1/ggml-small.bin',
      );
      return new Response('not a model');
    }, originalFetch);
    try {
      await assert.rejects(
        downloadSubtitleGenerationModel({
          config: DEFAULT_SUBTITLE_GENERATION_CONFIG,
          modelDirectory: directory,
        }),
        /integrity verification/,
      );
      assert.deepEqual(await readdir(directory), []);
    } finally {
      globalThis.fetch = originalFetch;
    }
  }));
