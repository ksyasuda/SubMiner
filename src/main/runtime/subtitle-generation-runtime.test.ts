import assert from 'node:assert/strict';
import test from 'node:test';
import { DEFAULT_SUBTITLE_GENERATION_CONFIG } from '../../shared/subtitle-generation';
import {
  createSubtitleGenerationRuntime,
  type SubtitleGenerationRuntimeDeps,
} from './subtitle-generation-runtime';

function fixture(overrides: Partial<SubtitleGenerationRuntimeDeps> = {}) {
  let mediaPath = '/video/episode.mkv';
  const commands: unknown[][] = [];
  const client = {
    connected: true,
    requestProperty: async (name: string): Promise<unknown> =>
      name === 'path' ? mediaPath : [{ type: 'audio', selected: true, 'ff-index': 3 }],
    request: async (command: unknown[]) => {
      commands.push(command);
      return { error: 'success' };
    },
  };
  const runtime = createSubtitleGenerationRuntime({
    getConfig: () => DEFAULT_SUBTITLE_GENERATION_CONFIG,
    getModelDirectory: () => '/models',
    getMpvClient: () => client,
    onProgress: () => {},
    resolveModel: async () => ({ kind: 'external', path: '/models/local.bin' }),
    generate: async () => '/video/episode.ja.generated.srt',
    ...overrides,
  });
  return {
    runtime,
    client,
    commands,
    changeMedia: () => {
      mediaPath = '/video/next.mkv';
    },
  };
}

test('generation uses the selected audio track and loads the timed SRT with zero delay', async () => {
  const { runtime, commands } = fixture({
    generate: async (input) => {
      assert.equal(input.mediaPath, '/video/episode.mkv');
      assert.equal(input.audioStreamIndex, 3);
      return '/video/generated.srt';
    },
  });
  assert.equal((await runtime.start()).ok, true);
  assert.deepEqual(commands, [
    ['sub-add', '/video/generated.srt', 'select', 'Generated Japanese', 'ja'],
    ['set_property', 'sub-delay', 0],
  ]);
});

test('generation preserves the output without attaching it to a different video', async () => {
  const subject = fixture({
    generate: async () => {
      subject.changeMedia();
      return '/video/generated.srt';
    },
  });
  const result = await subject.runtime.start();
  assert.equal(result.ok, true);
  assert.match(result.message, /Playback changed/);
  assert.deepEqual(subject.commands, []);
});

test('mpv load failure still reports where the generated subtitles were saved', async () => {
  const subject = fixture();
  subject.client.request = async () => ({ error: 'loading failed' });
  const result = await subject.runtime.start();
  assert.equal(result.ok, true);
  assert.match(result.message, /Subtitles saved:.*Could not finish loading/);
});

test('cancellation during the final media check keeps the saved file without loading it', async () => {
  const subject = fixture();
  const requestProperty = subject.client.requestProperty;
  let mediaChecks = 0;
  subject.client.requestProperty = async (name) => {
    if (name === 'path' && ++mediaChecks === 3) subject.runtime.cancel();
    return requestProperty(name);
  };
  const result = await subject.runtime.start();
  assert.equal(result.ok, true);
  assert.match(result.message, /Cancelled after saving/);
  assert.deepEqual(subject.commands, []);
});

test('only one job runs, cancellation reaches the worker, and status retains its result', async () => {
  let signal: AbortSignal | undefined;
  let entered = () => {};
  const started = new Promise<void>((resolve) => {
    entered = resolve;
  });
  const { runtime } = fixture({
    generate: async (input) => {
      signal = input.signal;
      input.onProgress?.({ stage: 'transcribe', percent: 25, message: 'Working' });
      entered();
      return new Promise((_, reject) =>
        input.signal?.addEventListener('abort', () => reject(new Error('Aborted')), { once: true }),
      );
    },
  });
  const first = runtime.start();
  await started;
  await assert.rejects(runtime.selectModel('medium'), /current operation/);
  assert.equal((await runtime.start()).ok, false);
  assert.equal((await runtime.download()).ok, false);
  const active = await runtime.getStatus();
  assert.equal(active.running, true);
  assert.equal(active.progress?.percent, 25);
  runtime.cancel();
  assert.equal(signal?.aborted, true);
  assert.deepEqual(await first, { ok: false, message: 'Cancelled.' });
  const completed = await runtime.getStatus();
  assert.equal(completed.running, false);
  assert.deepEqual(completed.lastResult, { ok: false, message: 'Cancelled.' });
});

test('external audio cannot silently generate from a different internal track', async () => {
  const subject = fixture({
    generate: async () => {
      assert.fail('must not transcribe');
    },
  });
  subject.client.requestProperty = async (name) =>
    name === 'path'
      ? '/video/episode.mkv'
      : [{ type: 'audio', selected: true, external: true, 'ff-index': 0 }];
  const result = await subject.runtime.start();
  assert.equal(result.ok, false);
  assert.match(result.message, /audio track inside/);
});

test('model selection is retained and used for status, download, and generation', async () => {
  const seen: string[] = [];
  const { runtime } = fixture({
    resolveModel: async (config) => ({
      kind: 'missing',
      path: `/models/${config.managedModel}.bin`,
    }),
    download: async ({ config }) => {
      seen.push(`download:${config.managedModel}`);
      return '/models/downloaded.bin';
    },
    generate: async ({ config }) => {
      seen.push(`generate:${config.managedModel}`);
      return '/video/generated.srt';
    },
  });
  const selected = await runtime.selectModel('medium');
  assert.equal(selected.managedModel, 'medium');
  assert.equal(selected.model.path, '/models/medium.bin');
  assert.equal((await runtime.getStatus()).managedModel, 'medium');
  assert.equal((await runtime.download()).ok, true);
  assert.equal((await runtime.start()).ok, true);
  assert.deepEqual(seen, ['download:medium', 'generate:medium']);
  assert.equal(DEFAULT_SUBTITLE_GENERATION_CONFIG.managedModel, 'small');
});

test('external model paths prevent managed selection, including unreadable overrides', async () => {
  const { runtime } = fixture({
    getConfig: () => ({
      ...DEFAULT_SUBTITLE_GENERATION_CONFIG,
      modelPath: '/missing/external.bin',
    }),
    resolveModel: async () => ({
      kind: 'invalid',
      path: '/missing/external.bin',
      message: 'Missing model',
    }),
  });
  assert.equal((await runtime.getStatus()).externalModelPath, '/missing/external.bin');
  await assert.rejects(runtime.selectModel('medium'), /Clear Model Path/);
});

test('speech detection is optional and downloading alone does not enable it', async () => {
  let installed = false;
  const paths: string[] = [];
  const { runtime } = fixture({
    resolveVadModel: async () => ({
      kind: installed ? 'managed' : 'missing',
      path: '/models/ggml-silero-v6.2.0.bin',
    }),
    downloadVad: async () => {
      installed = true;
      return '/models/ggml-silero-v6.2.0.bin';
    },
    generate: async ({ config }) => {
      paths.push(config.vadModelPath);
      return '/video/output.srt';
    },
  });
  assert.equal((await runtime.getStatus()).vad.enabled, false);
  assert.equal((await runtime.start()).ok, true);
  await runtime.setVadEnabled(true);
  assert.equal((await runtime.start()).ok, false);
  await runtime.setVadEnabled(false);
  assert.equal((await runtime.downloadVad()).ok, true);
  assert.equal((await runtime.getStatus()).vad.enabled, false);
  await runtime.setVadEnabled(true);
  assert.equal((await runtime.start()).ok, true);
  await runtime.setVadEnabled(false);
  assert.equal((await runtime.start()).ok, true);
  assert.deepEqual(paths, ['', '/models/ggml-silero-v6.2.0.bin', '']);
});

test('existing external speech model remains the default and survives session toggles', async () => {
  const { runtime } = fixture({
    getConfig: () => ({ ...DEFAULT_SUBTITLE_GENERATION_CONFIG, vadModelPath: '/external/vad.bin' }),
    resolveVadModel: async (config) => ({ kind: 'external', path: config.vadModelPath }),
    generate: async ({ config }) => {
      assert.equal(config.vadModelPath, '/external/vad.bin');
      return '/video/output.srt';
    },
  });
  assert.equal((await runtime.getStatus()).vad.enabled, true);
  await runtime.setVadEnabled(false);
  await runtime.setVadEnabled(true);
  assert.equal((await runtime.start()).ok, true);
});

test('speech model downloads share the job lock and cancellation', async () => {
  let enter = () => {};
  const started = new Promise<void>((resolve) => {
    enter = resolve;
  });
  const { runtime } = fixture({
    downloadVad: async ({ signal }) => {
      enter();
      return new Promise((_, reject) =>
        signal?.addEventListener('abort', () => reject(new Error('cancelled')), { once: true }),
      );
    },
  });
  const download = runtime.downloadVad();
  await started;
  assert.equal((await runtime.start()).ok, false);
  await assert.rejects(runtime.setVadEnabled(true), /current operation/);
  runtime.cancel();
  assert.deepEqual(await download, { ok: false, message: 'Cancelled.' });
});
