import test from 'node:test';
import assert from 'node:assert/strict';
import { createAniSkipRuntime, isRemoteMediaPath, AniSkipRuntimeDeps } from './aniskip-runtime';
import type { AniSkipMetadata } from './aniskip-metadata';

function readyMetadata(overrides: Partial<AniSkipMetadata> = {}): AniSkipMetadata {
  return {
    title: 'My Show',
    season: 1,
    episode: 1,
    source: 'fallback',
    malId: 1234,
    introStart: 10,
    introEnd: 95.5,
    lookupStatus: 'ready',
    ...overrides,
  };
}

function createHarness(options?: {
  enabled?: boolean;
  buttonKey?: string;
  metadata?: AniSkipMetadata | (() => Promise<AniSkipMetadata>);
  chapterList?: unknown;
  playbackFeedback?: boolean;
}) {
  const state = {
    enabled: options?.enabled ?? true,
    buttonKey: options?.buttonKey ?? 'TAB',
    commands: [] as unknown[][],
    osd: [] as string[],
    feedback: [] as string[],
    resolveCalls: [] as string[],
    connected: true,
    timePos: 0,
    chapterList: options?.chapterList ?? [],
  };

  const deps = {
    getAniSkipConfig: () => ({
      aniskipEnabled: state.enabled,
      aniskipButtonKey: state.buttonKey,
    }),
    resolveMetadataForFile: async (mediaPath) => {
      state.resolveCalls.push(mediaPath);
      const metadata = options?.metadata;
      if (typeof metadata === 'function') return metadata();
      return metadata ?? readyMetadata();
    },
    sendMpvCommand: (command) => {
      state.commands.push(command);
    },
    requestMpvProperty: async (name) => {
      if (name === 'chapter-list') return state.chapterList;
      return null;
    },
    isMpvConnected: () => state.connected,
    getCurrentTimePos: () => state.timePos,
    showMpvOsd: (text) => {
      state.osd.push(text);
    },
    ...(options?.playbackFeedback
      ? {
          showPlaybackFeedback: (text: string) => {
            state.feedback.push(text);
          },
        }
      : {}),
    logInfo: () => {},
    logWarn: () => {},
    logDebug: () => {},
  } satisfies AniSkipRuntimeDeps & { showPlaybackFeedback?: (text: string) => void };

  return { runtime: createAniSkipRuntime(deps), state };
}

function chapterListCommands(commands: unknown[][]): unknown[][] {
  return commands.filter(
    (command) => command[0] === 'set_property' && command[1] === 'chapter-list',
  );
}

async function flushAsync(): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, 0));
}

test('isRemoteMediaPath detects URLs but not local paths', () => {
  assert.equal(isRemoteMediaPath('https://example.com/stream.mkv'), true);
  assert.equal(isRemoteMediaPath('rtmp://example.com/live'), true);
  assert.equal(isRemoteMediaPath('/media/anime/show.mkv'), false);
  assert.equal(isRemoteMediaPath('C:\\media\\show.mkv'), false);
  assert.equal(isRemoteMediaPath(''), false);
});

test('media path change resolves metadata and sets AniSkip chapters', async () => {
  const { runtime, state } = createHarness({
    chapterList: [{ time: 0, title: 'Prologue' }],
  });

  runtime.handleMediaPathChange({ path: '/media/anime/My Show/ep1.mkv' });
  await flushAsync();

  assert.deepEqual(state.resolveCalls, ['/media/anime/My Show/ep1.mkv']);
  const chapterCommands = chapterListCommands(state.commands);
  assert.equal(chapterCommands.length, 1);
  const chapters = chapterCommands[0]![2] as Array<{ time: number; title: string }>;
  assert.deepEqual(chapters, [
    { time: 0, title: 'Prologue' },
    { time: 10, title: 'AniSkip Intro Start' },
    { time: 95.5, title: 'AniSkip Intro End' },
  ]);
  assert.deepEqual(runtime.getIntroWindow(), { start: 10, end: 95.5, malId: 1234 });
});

test('remote media paths and disabled config never resolve', async () => {
  const remote = createHarness();
  remote.runtime.handleMediaPathChange({ path: 'https://example.com/video.mkv' });
  await flushAsync();
  assert.deepEqual(remote.state.resolveCalls, []);

  const disabled = createHarness({ enabled: false });
  disabled.runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();
  assert.deepEqual(disabled.state.resolveCalls, []);
});

test('skip intro seeks to intro end only inside the intro window', async () => {
  const { runtime, state } = createHarness();
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();

  state.timePos = 200;
  runtime.handleClientMessage({ args: ['subminer-skip-intro'] });
  assert.deepEqual(state.osd, ['Skip intro only during intro']);

  state.timePos = 30;
  runtime.handleClientMessage({ args: ['subminer-skip-intro'] });
  assert.deepEqual(state.osd, ['Skip intro only during intro', 'Skipped intro']);
  const seek = state.commands.find(
    (command) => command[0] === 'set_property' && command[1] === 'time-pos',
  );
  assert.deepEqual(seek, ['set_property', 'time-pos', 95.5]);
});

test('skip intro reports unavailable when no window was found', () => {
  const { runtime, state } = createHarness();
  runtime.handleClientMessage({ args: ['subminer-skip-intro'] });
  assert.deepEqual(state.osd, ['Intro skip unavailable']);
});

test('time-pos prompt shows once near intro start', async () => {
  const { runtime, state } = createHarness({ buttonKey: 'TAB' });
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();

  runtime.handleTimePosChange({ time: 5 });
  assert.deepEqual(state.osd, []);

  runtime.handleTimePosChange({ time: 10.5 });
  runtime.handleTimePosChange({ time: 11 });
  assert.deepEqual(state.osd, ['You can skip by pressing TAB']);
});

test('prompt and skip messages use playback feedback when configured', async () => {
  const { runtime, state } = createHarness({ buttonKey: 'TAB', playbackFeedback: true });
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();

  runtime.handleTimePosChange({ time: 10.5 });
  state.timePos = 30;
  runtime.handleClientMessage({ args: ['subminer-skip-intro'] });

  assert.deepEqual(state.feedback, ['You can skip by pressing TAB', 'Skipped intro']);
  assert.deepEqual(state.osd, []);
});

test('connection change binds skip key and legacy fallback for custom keys', () => {
  const { runtime, state } = createHarness({ buttonKey: 'F6' });
  runtime.handleConnectionChange({ connected: true });
  assert.deepEqual(state.commands, [
    ['keybind', 'F6', 'script-message subminer-skip-intro'],
    ['keybind', 'y-k', 'script-message subminer-skip-intro'],
  ]);
});

test('default key binds without duplicate legacy fallback', () => {
  const { runtime, state } = createHarness({ buttonKey: 'TAB' });
  runtime.handleConnectionChange({ connected: true });
  assert.deepEqual(state.commands, [['keybind', 'TAB', 'script-message subminer-skip-intro']]);
});

test('config change rebinds key and disabling unbinds and clears chapters', async () => {
  const { runtime, state } = createHarness({ buttonKey: 'TAB' });
  runtime.handleConnectionChange({ connected: true });

  state.buttonKey = 'F6';
  runtime.applyConfigChange();
  assert.deepEqual(state.commands.slice(1), [
    ['keybind', 'TAB', ''],
    ['keybind', 'F6', 'script-message subminer-skip-intro'],
    ['keybind', 'y-k', 'script-message subminer-skip-intro'],
  ]);

  state.commands.length = 0;
  state.enabled = false;
  state.chapterList = [
    { time: 0, title: 'Prologue' },
    { time: 10, title: 'AniSkip Intro Start' },
    { time: 95.5, title: 'AniSkip Intro End' },
  ];
  runtime.applyConfigChange();
  await flushAsync();
  assert.deepEqual(state.commands, [
    ['keybind', 'F6', ''],
    ['keybind', 'y-k', ''],
    ['set_property', 'chapter-list', [{ time: 0, title: 'Prologue' }]],
  ]);
});

test('same-media reload re-applies chapters without a new lookup', async () => {
  const { runtime, state } = createHarness();
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();
  assert.equal(state.resolveCalls.length, 1);

  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();
  assert.equal(state.resolveCalls.length, 1);
  assert.equal(chapterListCommands(state.commands).length, 2);
});

test('aniskip refresh forces a fresh lookup for the current media', async () => {
  const { runtime, state } = createHarness();
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();
  assert.equal(state.resolveCalls.length, 1);

  runtime.handleClientMessage({ args: ['subminer-aniskip-refresh'] });
  await flushAsync();
  assert.equal(state.resolveCalls.length, 2);
});

test('media without an intro window is cached and never re-resolved on reload of another file', async () => {
  const { runtime, state } = createHarness({
    metadata: readyMetadata({
      malId: 1234,
      introStart: null,
      introEnd: null,
      lookupStatus: 'missing_payload',
    }),
  });
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();
  assert.equal(runtime.getIntroWindow(), null);
  assert.equal(chapterListCommands(state.commands).length, 0);

  runtime.handleMediaPathChange({ path: '/media/other.mkv' });
  await flushAsync();
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();
  assert.deepEqual(state.resolveCalls, ['/media/show.mkv', '/media/other.mkv']);
});

test('transient lookup failures are retried on the next media load', async () => {
  let failures = 0;
  const { runtime, state } = createHarness({
    metadata: async () => {
      failures += 1;
      return readyMetadata(
        failures === 1 ? { introStart: null, introEnd: null, lookupStatus: 'lookup_failed' } : {},
      );
    },
  });
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();
  assert.equal(runtime.getIntroWindow(), null);

  runtime.handleMediaPathChange({ path: '' });
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();
  assert.equal(state.resolveCalls.length, 2);
  assert.deepEqual(runtime.getIntroWindow(), { start: 10, end: 95.5, malId: 1234 });
});

test('disconnect clears bindings so reconnect rebinds the skip key', () => {
  const { runtime, state } = createHarness();
  runtime.handleConnectionChange({ connected: true });
  runtime.handleConnectionChange({ connected: false });
  runtime.handleConnectionChange({ connected: true });
  assert.deepEqual(state.commands, [
    ['keybind', 'TAB', 'script-message subminer-skip-intro'],
    ['keybind', 'TAB', 'script-message subminer-skip-intro'],
  ]);
});
