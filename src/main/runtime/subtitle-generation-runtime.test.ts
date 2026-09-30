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
    detectAcceleration: async () => ({ kind: 'unavailable' }),
    resolveModel: async () => ({ kind: 'external', path: '/models/local.bin' }),
    resolveTools: async (config) => ({
      ffmpeg: { kind: 'found', path: '/usr/bin/ffmpeg' },
      ffprobe: { kind: 'found', path: '/usr/bin/ffprobe' },
      whisper: { kind: 'found', path: '/usr/bin/whisper-cli' },
      vad: config.vadModelPath ? { kind: 'found', path: '/usr/bin/vad' } : null,
    }),
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

test('generation selects loaded dialogue references and excludes the signs track', async () => {
  const subject = fixture({
    generate: async (input) => {
      assert.deepEqual(input.references, [
        { label: 'English Full', delaySeconds: 0, source: { kind: 'embedded', streamIndex: 5 } },
      ]);
      return '/video/generated.srt';
    },
  });
  const request = subject.client.requestProperty;
  subject.client.requestProperty = async (name) =>
    name === 'track-list'
      ? [
          { type: 'audio', selected: true, 'ff-index': 3 },
          { type: 'sub', lang: 'eng', title: 'Signs & Songs', 'ff-index': 4 },
          { type: 'sub', lang: 'eng', title: 'English Full', 'ff-index': 5 },
        ]
      : request(name);
  assert.equal((await subject.runtime.start()).ok, true);
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

test('CUDA recommendations preserve selected and configured models and follow the Whisper path', async () => {
  let whisperPath = '/cuda/whisper-cli';
  const checked: string[] = [];
  const subject = fixture({
    getConfig: () => ({ ...DEFAULT_SUBTITLE_GENERATION_CONFIG, managedModel: 'medium' }),
    resolveTools: async () => ({
      ffmpeg: { kind: 'found', path: '/usr/bin/ffmpeg' },
      ffprobe: { kind: 'found', path: '/usr/bin/ffprobe' },
      whisper: { kind: 'found', path: whisperPath },
      vad: null,
    }),
    detectAcceleration: async (whisper) => {
      assert.equal(whisper.kind, 'found');
      if (whisper.kind !== 'found') throw new Error('Expected a Whisper executable');
      checked.push(whisper.path);
      return whisper.path.startsWith('/cuda/')
        ? { kind: 'nvidia-cuda', gpuName: 'NVIDIA Test GPU' }
        : { kind: 'unavailable' };
    },
  });
  const initial = await subject.runtime.getStatus();
  assert.deepEqual(initial.acceleration, { kind: 'nvidia-cuda', gpuName: 'NVIDIA Test GPU' });
  assert.equal(initial.managedModel, 'medium');
  const selected = await subject.runtime.selectModel('small');
  assert.equal(selected.managedModel, 'small');
  assert.equal(selected.acceleration.kind, 'nvidia-cuda');
  assert.deepEqual(checked, ['/cuda/whisper-cli']);
  whisperPath = '/cpu/whisper-cli';
  const changed = await subject.runtime.getStatus();
  assert.equal(changed.acceleration.kind, 'unavailable');
  assert.equal(changed.managedModel, 'small');
  assert.deepEqual(checked, ['/cuda/whisper-cli', '/cpu/whisper-cli']);
});

test('status reports the speech detector only while dialogue mode is on', async () => {
  const { runtime } = fixture({
    resolveVadModel: async () => ({ kind: 'managed', path: '/models/ggml-silero-v6.2.0.bin' }),
  });
  assert.equal((await runtime.getStatus()).tools.vad, null);
  await runtime.setVadEnabled(true);
  assert.deepEqual((await runtime.getStatus()).tools.vad, { kind: 'found', path: '/usr/bin/vad' });
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

test('a YouTube Whisper job shows in the modal status and returns its subtitle file', async () => {
  const { runtime } = fixture();
  let release = () => {};
  const job = runtime.generateYoutubeSubtitles({
    url: 'https://www.youtube.com/watch?v=abcdefghijk',
    signal: new AbortController().signal,
    generate: async (_signal, onProgress) => {
      onProgress({ stage: 'transcribe', percent: 40, message: 'Generating Japanese subtitles...' });
      await new Promise<void>((resolve) => {
        release = resolve;
      });
      return '/tmp/subs/youtube-whisper.ja.srt';
    },
  });

  const active = await runtime.getStatus();
  assert.equal(active.running, true);
  assert.equal(active.mediaPath, 'https://www.youtube.com/watch?v=abcdefghijk');
  assert.equal(active.progress?.percent, 40);
  assert.equal((await runtime.start()).ok, false);

  release();
  assert.equal(await job, '/tmp/subs/youtube-whisper.ja.srt');
  const done = await runtime.getStatus();
  assert.equal(done.running, false);
  assert.equal(done.mediaPath, '/video/episode.mkv');
});

test('cancelling a YouTube Whisper job resolves to null, and failures throw', async () => {
  const { runtime } = fixture();
  const cancelled = runtime.generateYoutubeSubtitles({
    url: 'https://www.youtube.com/watch?v=abcdefghijk',
    signal: new AbortController().signal,
    generate: (signal) =>
      new Promise((_, reject) =>
        signal.addEventListener('abort', () => reject(new Error('Aborted')), { once: true }),
      ),
  });
  runtime.cancel();
  assert.equal(await cancelled, null);
  assert.deepEqual((await runtime.getStatus()).lastResult, { ok: false, message: 'Cancelled.' });

  await assert.rejects(
    runtime.generateYoutubeSubtitles({
      url: 'https://www.youtube.com/watch?v=abcdefghijk',
      signal: new AbortController().signal,
      generate: async () => {
        throw new Error('No Whisper model found.');
      },
    }),
    /No Whisper model found/,
  );
});

function youtubeFixture(generateYoutube: SubtitleGenerationRuntimeDeps['generateYoutube']) {
  const player = {
    path: 'https://www.youtube.com/watch?v=abcdefghijk&pp=search',
    paused: false,
  };
  const commands: unknown[][] = [];
  const client = {
    connected: true,
    requestProperty: async (name: string): Promise<unknown> =>
      name === 'path' ? player.path : name === 'pause' ? player.paused : null,
    request: async (command: unknown[]) => {
      commands.push(command);
      if (command[0] === 'set_property' && command[1] === 'pause')
        player.paused = command[2] === true;
      return { error: 'success' };
    },
  };
  const { runtime } = fixture({ getMpvClient: () => client, generateYoutube });
  return { runtime, player, commands };
}

test('the modal generates for a playing YouTube video, pausing it until the subtitles load', async () => {
  let pausedDuringGeneration = false;
  let usedModel = '';
  const { runtime, player, commands } = youtubeFixture(async (input) => {
    pausedDuringGeneration = player.paused;
    usedModel = input.config.managedModel;
    assert.equal(input.url, 'https://www.youtube.com/watch?v=abcdefghijk');
    return '/tmp/subs/youtube-whisper.ja.srt';
  });

  assert.equal(
    (await runtime.getStatus()).mediaPath,
    'https://www.youtube.com/watch?v=abcdefghijk',
  );
  await runtime.selectModel('medium');
  const result = await runtime.start();

  assert.deepEqual(result, {
    ok: true,
    outputPath: '/tmp/subs/youtube-whisper.ja.srt',
    message: 'Japanese subtitles generated and loaded.',
  });
  assert.equal(pausedDuringGeneration, true);
  assert.equal(usedModel, 'medium');
  assert.deepEqual(
    commands.find((command) => command[0] === 'sub-add'),
    ['sub-add', '/tmp/subs/youtube-whisper.ja.srt', 'select', 'Generated Japanese', 'ja'],
  );
  assert.equal(player.paused, false);
});

test('switching videos cancels a YouTube job from the modal without loading its subtitles', async () => {
  let jobSignal: AbortSignal | null = null;
  const { runtime, player, commands } = youtubeFixture(
    (input) =>
      new Promise((_, reject) => {
        jobSignal = input.signal;
        input.signal.addEventListener('abort', () => reject(new Error('Aborted')), { once: true });
      }),
  );

  const job = runtime.start();
  while (!jobSignal) await new Promise((resolve) => setImmediate(resolve));
  runtime.handleMediaPathChange('https://www.youtube.com/watch?v=abcdefghijk');
  assert.equal((jobSignal as AbortSignal).aborted, false);
  player.path = 'https://www.youtube.com/watch?v=zyxwvutsrqp';
  runtime.handleMediaPathChange(player.path);

  assert.deepEqual(await job, { ok: false, message: 'Cancelled.' });
  assert.equal(
    commands.some((command) => command[0] === 'sub-add'),
    false,
  );
});
