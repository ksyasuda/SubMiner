import assert from 'node:assert/strict';
import test from 'node:test';
import type { MergedToken } from '../../types';
import {
  createAutoplayReadyGate,
  type AutoplayReadyGateDeps,
  type AutoplayReadySignal,
} from './autoplay-ready-gate';

type MpvCommand = Array<string | boolean>;

const PLUGIN_READY: MpvCommand = ['script-message', 'subminer-autoplay-ready'];
const FAKE_TIMER = 1 as unknown as ReturnType<typeof setTimeout>;

const settle = () => new Promise<void>((resolve) => setTimeout(resolve, 0));

const isUnpause = (command: MpvCommand): boolean =>
  command[0] === 'set_property' && command[1] === 'pause' && command[2] === false;

// Builds a gate against a connected fake mpv that starts paused. An unpause command
// flips `mpv.paused`; scheduled retries queue in `scheduled` until run explicitly.
function createGate(overrides: Partial<AutoplayReadyGateDeps> = {}) {
  const commands: MpvCommand[] = [];
  const scheduled: Array<() => void> = [];
  const mpv = { paused: true };

  const gate = createAutoplayReadyGate({
    isAppOwnedFlowInFlight: () => false,
    getCurrentMediaPath: () => '/media/video.mkv',
    getCurrentVideoPath: () => null,
    getPlaybackPaused: () => mpv.paused,
    getMpvClient: () => ({
      connected: true,
      requestProperty: async () => mpv.paused,
      send: ({ command }) => {
        commands.push(command);
        if (isUnpause(command)) mpv.paused = false;
      },
    }),
    signalPluginAutoplayReady: () => {
      commands.push(PLUGIN_READY);
    },
    schedule: (callback) => {
      scheduled.push(callback);
      return FAKE_TIMER;
    },
    logDebug: () => {},
    ...overrides,
  });

  return {
    gate,
    commands,
    scheduled,
    mpv,
    pluginSignals: () => commands.filter((command) => command[0] === 'script-message'),
    unpauseCount: () => commands.filter(isUnpause).length,
    // Delivers a forced-while-paused subtitle signal and lets async release attempts settle.
    signal: async (text: string, tokens: MergedToken[] | null = null) => {
      gate.maybeSignalPluginAutoplayReady({ text, tokens }, { forceWhilePaused: true });
      await settle();
    },
    // Runs queued retries one at a time, including retries they schedule, until none remain.
    runScheduled: async () => {
      for (let callback = scheduled.shift(); callback; callback = scheduled.shift()) {
        callback();
        await settle();
      }
    },
  };
}

test('autoplay ready gate signals the plugin once per media across duplicate signals and release retries', async () => {
  const h = createGate();

  await h.signal('字幕');
  await h.signal('字幕');
  await h.runScheduled();

  assert.deepEqual(h.pluginSignals(), [PLUGIN_READY]);
  assert.equal(h.unpauseCount(), 1);
});

test('autoplay ready gate requests overlay pointer recovery once per released media', async () => {
  let pointerRecoveryRequests = 0;
  const h = createGate({
    requestOverlayPointerRecovery: () => {
      pointerRecoveryRequests += 1;
    },
  });

  await h.signal('字幕');
  await h.signal('字幕その2');

  assert.equal(pointerRecoveryRequests, 1);
});

test('autoplay ready gate reports the released autoplay signal once', async () => {
  const releasedSignals: string[] = [];
  const h = createGate({
    onAutoplayReadyReleased: (signal) => {
      releasedSignals.push(signal.payload.text);
    },
  });

  await h.signal('__warm__');
  await h.signal('次の字幕');

  assert.deepEqual(releasedSignals, ['__warm__']);
});

const manualPauseCases: Array<{
  name: string;
  afterManualPause: (h: ReturnType<typeof createGate>) => Promise<void>;
}> = [
  {
    name: 'autoplay ready gate does not unpause again after a later manual pause on the same media',
    afterManualPause: (h) => h.signal('字幕その2'),
  },
  {
    name: 'autoplay ready gate cancels release retries after playback is paused again',
    afterManualPause: (h) => h.runScheduled(),
  },
];

for (const c of manualPauseCases) {
  test(c.name, async () => {
    const h = createGate();

    await h.signal('字幕');
    assert.equal(h.unpauseCount(), 1);

    h.mpv.paused = true;
    await c.afterManualPause(h);

    assert.equal(h.unpauseCount(), 1);
    assert.equal(h.mpv.paused, true);
  });
}

test('autoplay ready gate suppresses release after manual current-media dismissal', async () => {
  const h = createGate();

  h.gate.markCurrentMediaAutoplayReady();
  await h.signal('字幕');

  assert.deepEqual(h.commands, []);
});

const deferredReleaseCases: Array<{
  name: string;
  release: (h: ReturnType<typeof createGate>) => Promise<void>;
}> = [
  {
    name: 'autoplay ready gate defers plugin readiness until the signal target is ready',
    release: async (h) => {
      h.gate.flushPendingAutoplayReadySignal();
      await settle();
    },
  },
  {
    name: 'autoplay ready gate retries deferred readiness without an external flush event',
    release: (h) => h.runScheduled(),
  },
];

for (const c of deferredReleaseCases) {
  test(c.name, async () => {
    let targetReady = false;
    const h = createGate({ isSignalTargetReady: () => targetReady });

    await h.signal('字幕');
    assert.deepEqual(h.commands, []);
    assert.equal(h.scheduled.length, 1);

    targetReady = true;
    await c.release(h);

    assert.deepEqual(h.pluginSignals(), [PLUGIN_READY]);
    assert.equal(h.unpauseCount(), 1);
  });
}

test('autoplay ready gate keeps deferred startup readiness retries active for cold starts', async () => {
  const h = createGate({ isSignalTargetReady: () => false });

  await h.signal('__warm__');

  for (let attempt = 1; attempt <= 100; attempt += 1) {
    assert.equal(h.scheduled.length, 1, `missing deferred readiness retry ${attempt}`);
    h.scheduled.shift()?.();
    await settle();
  }

  assert.deepEqual(h.commands, []);
});

test('autoplay ready gate drops deferred readiness after media changes before flush', async () => {
  let targetReady = false;
  let currentMediaPath = '/media/video-1.mkv';
  const h = createGate({
    getCurrentMediaPath: () => currentMediaPath,
    isSignalTargetReady: () => targetReady,
  });

  await h.signal('字幕');

  currentMediaPath = '/media/video-2.mkv';
  targetReady = true;
  h.gate.flushPendingAutoplayReadySignal();
  await settle();

  assert.deepEqual(h.commands, []);
});

test('autoplay ready gate passes the pending subtitle signal to the readiness predicate', async () => {
  const observed: Array<Pick<AutoplayReadySignal, 'requestedAtMs'> & { text: string }> = [];
  let targetReadyText: string | null = null;
  let now = 1_000;
  const h = createGate({
    isSignalTargetReady: (signal) => {
      observed.push({ text: signal.payload.text, requestedAtMs: signal.requestedAtMs });
      return targetReadyText === signal.payload.text;
    },
    now: () => now,
  });

  await h.signal('字幕');

  assert.deepEqual(observed, [{ text: '字幕', requestedAtMs: 1_000 }]);
  assert.deepEqual(h.commands, []);

  now = 2_000;
  targetReadyText = '字幕';
  h.gate.flushPendingAutoplayReadySignal();
  await settle();

  // The flushed signal keeps its original request time instead of the flush time.
  assert.deepEqual(observed.at(-1), { text: '字幕', requestedAtMs: 1_000 });
  assert.deepEqual(h.pluginSignals(), [PLUGIN_READY]);
});

test('autoplay ready gate ignores untokenized signals until tokenization warmup is ready', async () => {
  let tokenizationReady = false;
  const h = createGate({ isTokenizationReady: () => tokenizationReady });

  await h.signal('字幕');
  assert.deepEqual(h.commands, []);

  tokenizationReady = true;
  await h.signal('字幕');

  assert.deepEqual(h.pluginSignals(), [PLUGIN_READY]);
});

test('autoplay ready gate releases tokenized signals while tokenization warmup is pending', async () => {
  const h = createGate({ isTokenizationReady: () => false });

  await h.signal('字幕', [{ surface: '字幕' }] as unknown as MergedToken[]);

  assert.deepEqual(h.pluginSignals(), [PLUGIN_READY]);
});
