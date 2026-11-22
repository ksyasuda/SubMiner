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
    intro: { start: 10, end: 95.5 },
    outro: null,
    lookupStatus: 'ready',
    ...overrides,
  };
}

const NO_SKIP_TIMES = { intro: null, outro: null } as const;

function createHarness(options?: {
  enabled?: boolean;
  buttonKey?: string;
  metadata?: AniSkipMetadata | (() => Promise<AniSkipMetadata>);
  chapterList?: unknown;
  playbackFeedback?: boolean;
  duration?: number | null;
}) {
  const state = {
    enabled: options?.enabled ?? true,
    buttonKey: options?.buttonKey ?? 'TAB',
    commands: [] as unknown[][],
    osd: [] as string[],
    feedback: [] as string[],
    resolveCalls: [] as string[],
    episodeLengths: [] as number[],
    connected: true,
    duration: options?.duration === undefined ? 1440 : options.duration,
    timePos: 0,
    chapterList: options?.chapterList ?? [],
  };

  const deps = {
    getAniSkipConfig: () => ({
      aniskipEnabled: state.enabled,
      aniskipButtonKey: state.buttonKey,
    }),
    resolveMetadataForFile: async (mediaPath, episodeLength) => {
      state.resolveCalls.push(mediaPath);
      state.episodeLengths.push(episodeLength);
      const metadata = options?.metadata;
      if (typeof metadata === 'function') return metadata();
      return metadata ?? readyMetadata();
    },
    sendMpvCommand: (command) => {
      state.commands.push(command);
    },
    requestMpvProperty: async (name) => {
      if (name === 'chapter-list') return state.chapterList;
      if (name === 'duration') {
        if (state.duration === null) throw new Error('property unavailable');
        return state.duration;
      }
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

function seekCommands(commands: unknown[][]): unknown[][] {
  return commands.filter((command) => command[0] === 'set_property' && command[1] === 'time-pos');
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
  assert.deepEqual(state.episodeLengths, [1440]);
  const chapterCommands = chapterListCommands(state.commands);
  assert.equal(chapterCommands.length, 1);
  const chapters = chapterCommands[0]![2] as Array<{ time: number; title: string }>;
  assert.deepEqual(chapters, [
    { time: 0, title: 'Prologue' },
    { time: 10, title: 'AniSkip Intro Start' },
    { time: 95.5, title: 'AniSkip Intro End' },
  ]);
  assert.deepEqual(runtime.getSkipWindows(), {
    intro: { start: 10, end: 95.5, source: 'aniskip' },
    outro: null,
  });
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
  assert.deepEqual(state.osd, ['Skip only during intro or outro']);

  state.timePos = 30;
  runtime.handleClientMessage({ args: ['subminer-skip-intro'] });
  assert.deepEqual(state.osd, ['Skip only during intro or outro', 'Skipped intro']);
  assert.deepEqual(seekCommands(state.commands), [['set_property', 'time-pos', 95.5]]);
});

test('skip intro reports unavailable when no window was found', () => {
  const { runtime, state } = createHarness();
  runtime.handleClientMessage({ args: ['subminer-skip-intro'] });
  assert.deepEqual(state.osd, ['Intro/outro skip unavailable']);
});

test('time-pos prompt shows once near intro start', async () => {
  const { runtime, state } = createHarness({ buttonKey: 'TAB' });
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();

  runtime.handleTimePosChange({ time: 5 });
  assert.deepEqual(state.osd, []);

  runtime.handleTimePosChange({ time: 10.5 });
  runtime.handleTimePosChange({ time: 11 });
  assert.deepEqual(state.osd, ['You can skip the intro by pressing TAB']);
});

test('prompt and skip messages use playback feedback when configured', async () => {
  const { runtime, state } = createHarness({ buttonKey: 'TAB', playbackFeedback: true });
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();

  runtime.handleTimePosChange({ time: 10.5 });
  state.timePos = 30;
  runtime.handleClientMessage({ args: ['subminer-skip-intro'] });

  assert.deepEqual(state.feedback, ['You can skip the intro by pressing TAB', 'Skipped intro']);
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

test('definite misses are cached and not looked up again for the same file', async () => {
  const { runtime, state } = createHarness({
    metadata: readyMetadata({ ...NO_SKIP_TIMES, lookupStatus: 'missing_payload' }),
  });
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();
  assert.deepEqual(runtime.getSkipWindows(), NO_SKIP_TIMES);
  assert.equal(chapterListCommands(state.commands).length, 0);

  runtime.handleMediaPathChange({ path: '/media/other.mkv' });
  await flushAsync();
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();
  assert.deepEqual(state.resolveCalls, ['/media/show.mkv', '/media/other.mkv']);
});

test('transient lookup failures are not cached and retry on the next media load', async () => {
  let calls = 0;
  const { runtime, state } = createHarness({
    metadata: async () => {
      calls += 1;
      return readyMetadata(calls === 1 ? { ...NO_SKIP_TIMES, lookupStatus: 'lookup_failed' } : {});
    },
  });
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();
  assert.deepEqual(runtime.getSkipWindows(), NO_SKIP_TIMES);

  runtime.handleMediaPathChange({ path: '' });
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();
  assert.equal(state.resolveCalls.length, 2);
  assert.deepEqual(runtime.getSkipWindows().intro, { start: 10, end: 95.5, source: 'aniskip' });
});

test('resolution waits for a known duration before reading chapters or AniSkip', async () => {
  const { runtime, state } = createHarness({ duration: null });
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();
  assert.deepEqual(state.resolveCalls, []);

  runtime.handleDurationChange({ duration: 1425.2 });
  await flushAsync();
  assert.deepEqual(state.episodeLengths, [1425.2]);
});

test('file chapters with both an opening and an ending skip the AniSkip lookup', async () => {
  const { runtime, state } = createHarness({
    chapterList: [
      { time: 0, title: 'Prologue' },
      { time: 120, title: 'Opening' },
      { time: 210, title: 'Part A' },
      { time: 1300, title: 'Ending' },
      { time: 1390, title: 'Preview' },
    ],
  });
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();

  assert.deepEqual(state.resolveCalls, []);
  assert.deepEqual(runtime.getSkipWindows(), {
    intro: { start: 120, end: 210, source: 'chapters' },
    outro: { start: 1300, end: 1390, source: 'chapters' },
  });
  // The file already has chapter boundaries there, so no AniSkip markers are added.
  assert.equal(chapterListCommands(state.commands).length, 0);
});

test('AniSkip fills in the window the file chapters lack', async () => {
  const { runtime, state } = createHarness({
    chapterList: [
      { time: 0, title: 'Opening' },
      { time: 90, title: 'Episode' },
    ],
    metadata: readyMetadata({
      intro: { start: 3, end: 93 },
      outro: { start: 1330, end: 1420 },
    }),
  });
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();

  assert.equal(state.resolveCalls.length, 1);
  assert.deepEqual(runtime.getSkipWindows(), {
    intro: { start: 0, end: 90, source: 'chapters' },
    outro: { start: 1330, end: 1420, source: 'aniskip' },
  });
  assert.deepEqual(chapterListCommands(state.commands)[0]![2], [
    { time: 0, title: 'Opening' },
    { time: 90, title: 'Episode' },
    { time: 1330, title: 'AniSkip Outro Start' },
    { time: 1420, title: 'AniSkip Outro End' },
  ]);
});

test('remote streams use file chapters without an AniSkip lookup', async () => {
  const { runtime, state } = createHarness({
    chapterList: [
      { time: 0, title: 'Intro' },
      { time: 30, title: 'OP' },
      { time: 120, title: 'Episode' },
    ],
  });
  runtime.handleMediaPathChange({ path: 'https://jellyfin.example/Videos/1/stream.mkv' });
  await flushAsync();

  assert.deepEqual(state.resolveCalls, []);
  assert.deepEqual(runtime.getSkipWindows().intro, { start: 30, end: 120, source: 'chapters' });
});

test('skip key skips the outro inside its window and prompts at its start', async () => {
  const { runtime, state } = createHarness({
    metadata: readyMetadata({ outro: { start: 1330, end: 1420 } }),
  });
  runtime.handleMediaPathChange({ path: '/media/show.mkv' });
  await flushAsync();

  runtime.handleTimePosChange({ time: 1331 });
  runtime.handleTimePosChange({ time: 1332 });
  assert.deepEqual(state.osd, ['You can skip the outro by pressing TAB']);

  state.timePos = 1329.8;
  runtime.handleClientMessage({ args: ['subminer-skip-intro'] });
  state.timePos = 1425;
  runtime.handleClientMessage({ args: ['subminer-skip-intro'] });
  assert.deepEqual(seekCommands(state.commands), [['set_property', 'time-pos', 1420]]);
  assert.deepEqual(state.osd.slice(1), ['Skipped outro', 'Skip only during intro or outro']);
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
