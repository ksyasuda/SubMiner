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
    session: { directory: async () => '/session', dispose: async () => {} },
    getConfig: () => DEFAULT_SUBTITLE_GENERATION_CONFIG,
    getModelDirectory: () => '/models',
    getCacheDirectory: () => '/cache/generated-subtitles',
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

test('stream alternatives are snapshotted without changing the playback identity', async () => {
  const url = 'http://127.0.0.1:7777/video/high';
  const alternatives = [
    {
      kind: 'audio' as const,
      url: 'http://127.0.0.1:7777/audio/ja',
      label: 'Japanese audio',
      httpHeaders: { headers: {}, userAgent: null },
    },
  ];
  const subject = fixture({
    getAlternativeSources: (mediaPath) => {
      assert.equal(mediaPath, url);
      return alternatives;
    },
    generate: async (input) => {
      assert.equal(input.mediaPath, url);
      assert.deepEqual(input.remote?.alternatives, alternatives);
      return '/cache/generated.srt';
    },
  });
  const original = subject.client.requestProperty;
  subject.client.requestProperty = async (name) => (name === 'path' ? url : original(name));
  assert.equal((await subject.runtime.start()).ok, true);
  assert.deepEqual(
    subject.commands.map((command) => command[0]),
    ['sub-add', 'set_property'],
  );
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

test('stream generation snapshots mpv headers and saves in the cache before loading', async () => {
  const url = 'http://127.0.0.1:7777/proxy/episode.m3u8';
  const subject = fixture({
    generate: async (input) => {
      assert.equal(input.mediaPath, url);
      assert.equal(input.audioStreamIndex, 3);
      assert.deepEqual(input.remote, {
        cacheDirectory: '/cache/generated-subtitles',
        sessionDirectory: '/session',
        httpHeaders: {
          headers: { Referer: 'https://anime.example/', 'X-Stream': 'episode' },
          userAgent: 'Anime Player',
        },
      });
      return '/cache/generated-subtitles/episode.ja.generated.srt';
    },
  });
  const request = subject.client.requestProperty;
  subject.client.requestProperty = async (name) => {
    if (name === 'path') return url;
    if (name === 'file-local-options/http-header-fields')
      return ['Referer: https://anime.example/', 'X-Stream: episode'];
    if (name === 'file-local-options/user-agent') return 'Anime Player';
    return request(name);
  };
  assert.equal((await subject.runtime.getStatus()).mediaPath, url);
  assert.equal((await subject.runtime.start()).ok, true);
  assert.deepEqual(subject.commands[0], [
    'sub-add',
    '/cache/generated-subtitles/episode.ja.generated.srt',
    'select',
    'Generated Japanese',
    'ja',
  ]);
});

test('stream changes during header capture stop generation; changes during transcription keep the saved result', async () => {
  for (const duringCapture of [true, false]) {
    let current = 'https://anime.example/episode.m3u8';
    const next = 'https://anime.example/next.m3u8';
    let generated = false;
    const subject = fixture({
      generate: async () => {
        generated = true;
        current = next;
        return '/cache/generated.srt';
      },
    });
    const request = subject.client.requestProperty;
    subject.client.requestProperty = async (name) => {
      if (name === 'path') return current;
      if (duringCapture && name === 'file-local-options/http-header-fields') current = next;
      return request(name);
    };
    const result = await subject.runtime.start();
    assert.equal(result.ok, !duringCapture);
    assert.equal(generated, !duringCapture);
    assert.match(result.message, /changed/);
    assert.deepEqual(subject.commands, []);
  }
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

test('a selected HTTP audio track keeps its own stream index, headers and playback delay', async () => {
  const mediaPath = 'https://example.test/video';
  const audioUrl = 'https://example.test/japanese.m4a';
  const httpHeaders = { headers: { Referer: 'https://audio.example/' }, userAgent: null };
  const subject = fixture({
    getAlternativeSources: () => [
      { kind: 'audio', url: audioUrl, label: 'Japanese audio', httpHeaders },
    ],
    generate: async (input) => {
      assert.equal(input.mediaPath, mediaPath);
      assert.equal(input.audioStreamIndex, undefined);
      assert.deepEqual(input.remote?.selectedAudio, {
        url: audioUrl,
        audioStreamIndex: 0,
        delaySeconds: 0.25,
        httpHeaders,
      });
      return '/cache/generated.srt';
    },
  });
  subject.client.requestProperty = async (name) => {
    if (name === 'path') return mediaPath;
    if (name === 'audio-delay') return 0.25;
    if (name === 'track-list')
      return [
        {
          type: 'audio',
          selected: true,
          external: true,
          'external-filename': audioUrl,
          'ff-index': 0,
        },
      ];
    return null;
  };
  assert.equal((await subject.runtime.start()).ok, true);
  assert.equal(subject.commands[0]?.[0], 'sub-add');
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

test('shutdown cancels and drains the active job before deleting session files', async () => {
  const events: string[] = [];
  let started = () => {};
  const ready = new Promise<void>((resolve) => {
    started = resolve;
  });
  const { runtime } = fixture({
    session: {
      directory: async () => '/session',
      dispose: async () => {
        events.push('disposed');
      },
    },
    generate: async (input) => {
      started();
      await new Promise<void>((resolve) =>
        input.signal?.addEventListener(
          'abort',
          () => {
            events.push('aborted');
            setImmediate(() => {
              events.push('job-cleaned');
              resolve();
            });
          },
          { once: true },
        ),
      );
      throw new Error('cancelled');
    },
  });
  const job = runtime.start();
  await ready;
  await runtime.dispose();
  assert.equal((await job).ok, false);
  assert.deepEqual(events, ['aborted', 'job-cleaned', 'disposed']);
  assert.equal((await runtime.start()).ok, false);
});
