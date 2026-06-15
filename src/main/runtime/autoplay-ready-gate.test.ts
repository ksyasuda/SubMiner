import assert from 'node:assert/strict';
import test from 'node:test';
import { createAutoplayReadyGate } from './autoplay-ready-gate';

test('autoplay ready gate suppresses duplicate media signals for the same media', async () => {
  const commands: Array<Array<string | boolean>> = [];
  const scheduled: Array<() => void> = [];

  const gate = createAutoplayReadyGate({
    isAppOwnedFlowInFlight: () => false,
    getCurrentMediaPath: () => '/media/video.mkv',
    getCurrentVideoPath: () => null,
    getPlaybackPaused: () => true,
    getMpvClient: () =>
      ({
        connected: true,
        requestProperty: async () => true,
        send: ({ command }: { command: Array<string | boolean> }) => {
          commands.push(command);
        },
      }) as never,
    signalPluginAutoplayReady: () => {
      commands.push(['script-message', 'subminer-autoplay-ready']);
    },
    schedule: (callback) => {
      scheduled.push(callback);
      return 1 as never;
    },
    logDebug: () => {},
  });

  gate.maybeSignalPluginAutoplayReady({ text: '字幕', tokens: null });
  gate.maybeSignalPluginAutoplayReady({ text: '字幕', tokens: null });

  await new Promise((resolve) => setTimeout(resolve, 0));
  const firstScheduled = scheduled.shift();
  firstScheduled?.();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(
    commands.filter((command) => command[0] === 'script-message'),
    [['script-message', 'subminer-autoplay-ready']],
  );
  assert.ok(
    commands.some(
      (command) => command[0] === 'set_property' && command[1] === 'pause' && command[2] === false,
    ),
  );
});

test('autoplay ready gate retry loop does not re-signal plugin readiness', async () => {
  const commands: Array<Array<string | boolean>> = [];
  const scheduled: Array<() => void> = [];

  const gate = createAutoplayReadyGate({
    isAppOwnedFlowInFlight: () => false,
    getCurrentMediaPath: () => '/media/video.mkv',
    getCurrentVideoPath: () => null,
    getPlaybackPaused: () => true,
    getMpvClient: () =>
      ({
        connected: true,
        requestProperty: async () => true,
        send: ({ command }: { command: Array<string | boolean> }) => {
          commands.push(command);
        },
      }) as never,
    signalPluginAutoplayReady: () => {
      commands.push(['script-message', 'subminer-autoplay-ready']);
    },
    schedule: (callback) => {
      scheduled.push(callback);
      return 1 as never;
    },
    logDebug: () => {},
  });

  gate.maybeSignalPluginAutoplayReady({ text: '字幕', tokens: null }, { forceWhilePaused: true });

  await new Promise((resolve) => setTimeout(resolve, 0));
  for (const callback of scheduled.splice(0, 3)) {
    callback();
    await new Promise((resolve) => setTimeout(resolve, 0));
  }

  assert.deepEqual(
    commands.filter((command) => command[0] === 'script-message'),
    [['script-message', 'subminer-autoplay-ready']],
  );
  assert.equal(
    commands.filter(
      (command) => command[0] === 'set_property' && command[1] === 'pause' && command[2] === false,
    ).length > 0,
    true,
  );
});

test('autoplay ready gate requests overlay pointer recovery when media readiness is signaled', async () => {
  const commands: Array<Array<string | boolean>> = [];
  let pointerRecoveryRequests = 0;

  const gate = createAutoplayReadyGate({
    isAppOwnedFlowInFlight: () => false,
    getCurrentMediaPath: () => '/media/video.mkv',
    getCurrentVideoPath: () => null,
    getPlaybackPaused: () => true,
    getMpvClient: () =>
      ({
        connected: true,
        requestProperty: async () => true,
        send: ({ command }: { command: Array<string | boolean> }) => {
          commands.push(command);
        },
      }) as never,
    signalPluginAutoplayReady: () => {
      commands.push(['script-message', 'subminer-autoplay-ready']);
    },
    requestOverlayPointerRecovery: () => {
      pointerRecoveryRequests += 1;
    },
    schedule: (callback) => {
      queueMicrotask(callback);
      return 1 as never;
    },
    logDebug: () => {},
  });

  gate.maybeSignalPluginAutoplayReady({ text: '字幕', tokens: null }, { forceWhilePaused: true });
  await new Promise((resolve) => setTimeout(resolve, 0));

  gate.maybeSignalPluginAutoplayReady(
    { text: '字幕その2', tokens: null },
    { forceWhilePaused: true },
  );
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.equal(pointerRecoveryRequests, 1);
});

test('autoplay ready gate reports the released autoplay signal once', async () => {
  const releasedSignals: string[] = [];

  const gate = createAutoplayReadyGate({
    isAppOwnedFlowInFlight: () => false,
    getCurrentMediaPath: () => '/media/video.mkv',
    getCurrentVideoPath: () => null,
    getPlaybackPaused: () => true,
    getMpvClient: () =>
      ({
        connected: true,
        requestProperty: async () => true,
        send: () => {},
      }) as never,
    signalPluginAutoplayReady: () => {},
    onAutoplayReadyReleased: (signal) => {
      releasedSignals.push(signal.payload.text);
    },
    schedule: (callback) => {
      queueMicrotask(callback);
      return 1 as never;
    },
    logDebug: () => {},
  });

  gate.maybeSignalPluginAutoplayReady(
    { text: '__warm__', tokens: null },
    { forceWhilePaused: true },
  );
  await new Promise((resolve) => setTimeout(resolve, 0));

  gate.maybeSignalPluginAutoplayReady(
    { text: '次の字幕', tokens: null },
    { forceWhilePaused: true },
  );
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(releasedSignals, ['__warm__']);
});

test('autoplay ready gate does not unpause again after a later manual pause on the same media', async () => {
  const commands: Array<Array<string | boolean>> = [];
  let playbackPaused = true;

  const gate = createAutoplayReadyGate({
    isAppOwnedFlowInFlight: () => false,
    getCurrentMediaPath: () => '/media/video.mkv',
    getCurrentVideoPath: () => null,
    getPlaybackPaused: () => playbackPaused,
    getMpvClient: () =>
      ({
        connected: true,
        requestProperty: async () => playbackPaused,
        send: ({ command }: { command: Array<string | boolean> }) => {
          commands.push(command);
          if (command[0] === 'set_property' && command[1] === 'pause' && command[2] === false) {
            playbackPaused = false;
          }
        },
      }) as never,
    signalPluginAutoplayReady: () => {
      commands.push(['script-message', 'subminer-autoplay-ready']);
    },
    schedule: (callback) => {
      queueMicrotask(callback);
      return 1 as never;
    },
    logDebug: () => {},
  });

  gate.maybeSignalPluginAutoplayReady({ text: '字幕', tokens: null }, { forceWhilePaused: true });
  await new Promise((resolve) => setTimeout(resolve, 0));

  playbackPaused = true;
  gate.maybeSignalPluginAutoplayReady(
    { text: '字幕その2', tokens: null },
    { forceWhilePaused: true },
  );
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.equal(
    commands.filter(
      (command) => command[0] === 'set_property' && command[1] === 'pause' && command[2] === false,
    ).length,
    1,
  );
});

test('autoplay ready gate cancels release retries after playback is paused again', async () => {
  const commands: Array<Array<string | boolean>> = [];
  const scheduled: Array<() => void> = [];
  let playbackPaused = true;

  const gate = createAutoplayReadyGate({
    isAppOwnedFlowInFlight: () => false,
    getCurrentMediaPath: () => '/media/video.mkv',
    getCurrentVideoPath: () => null,
    getPlaybackPaused: () => playbackPaused,
    getMpvClient: () =>
      ({
        connected: true,
        requestProperty: async () => playbackPaused,
        send: ({ command }: { command: Array<string | boolean> }) => {
          commands.push(command);
          if (command[0] === 'set_property' && command[1] === 'pause' && command[2] === false) {
            playbackPaused = false;
          }
        },
      }) as never,
    signalPluginAutoplayReady: () => {
      commands.push(['script-message', 'subminer-autoplay-ready']);
    },
    schedule: (callback) => {
      scheduled.push(callback);
      return 1 as never;
    },
    logDebug: () => {},
  });

  gate.maybeSignalPluginAutoplayReady({ text: '字幕', tokens: null }, { forceWhilePaused: true });
  await new Promise((resolve) => setTimeout(resolve, 0));

  playbackPaused = true;
  const retry = scheduled.shift();
  retry?.();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.equal(
    commands.filter(
      (command) => command[0] === 'set_property' && command[1] === 'pause' && command[2] === false,
    ).length,
    1,
  );
});

test('autoplay ready gate suppresses release after manual current-media dismissal', async () => {
  const commands: Array<Array<string | boolean>> = [];

  const gate = createAutoplayReadyGate({
    isAppOwnedFlowInFlight: () => false,
    getCurrentMediaPath: () => '/media/video.mkv',
    getCurrentVideoPath: () => null,
    getPlaybackPaused: () => true,
    getMpvClient: () =>
      ({
        connected: true,
        requestProperty: async () => true,
        send: ({ command }: { command: Array<string | boolean> }) => {
          commands.push(command);
        },
      }) as never,
    signalPluginAutoplayReady: () => {
      commands.push(['script-message', 'subminer-autoplay-ready']);
    },
    schedule: (callback) => {
      queueMicrotask(callback);
      return 1 as never;
    },
    logDebug: () => {},
  });

  gate.markCurrentMediaAutoplayReady();
  gate.maybeSignalPluginAutoplayReady({ text: '字幕', tokens: null }, { forceWhilePaused: true });
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(commands, []);
});

test('autoplay ready gate defers plugin readiness until the signal target is ready', async () => {
  const commands: Array<Array<string | boolean>> = [];
  let targetReady = false;

  const gate = createAutoplayReadyGate({
    isAppOwnedFlowInFlight: () => false,
    getCurrentMediaPath: () => '/media/video.mkv',
    getCurrentVideoPath: () => null,
    getPlaybackPaused: () => true,
    getMpvClient: () =>
      ({
        connected: true,
        requestProperty: async () => true,
        send: ({ command }: { command: Array<string | boolean> }) => {
          commands.push(command);
        },
      }) as never,
    signalPluginAutoplayReady: () => {
      commands.push(['script-message', 'subminer-autoplay-ready']);
    },
    isSignalTargetReady: () => targetReady,
    schedule: (callback) => {
      queueMicrotask(callback);
      return 1 as never;
    },
    logDebug: () => {},
  });

  gate.maybeSignalPluginAutoplayReady({ text: '字幕', tokens: null }, { forceWhilePaused: true });
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(commands, []);

  targetReady = true;
  gate.flushPendingAutoplayReadySignal();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(
    commands.filter((command) => command[0] === 'script-message'),
    [['script-message', 'subminer-autoplay-ready']],
  );
  assert.equal(
    commands.some(
      (command) => command[0] === 'set_property' && command[1] === 'pause' && command[2] === false,
    ),
    true,
  );
});

test('autoplay ready gate retries deferred readiness without an external flush event', async () => {
  const commands: Array<Array<string | boolean>> = [];
  const scheduled: Array<() => void> = [];
  let targetReady = false;

  const gate = createAutoplayReadyGate({
    isAppOwnedFlowInFlight: () => false,
    getCurrentMediaPath: () => '/media/video.mkv',
    getCurrentVideoPath: () => null,
    getPlaybackPaused: () => true,
    getMpvClient: () =>
      ({
        connected: true,
        requestProperty: async () => true,
        send: ({ command }: { command: Array<string | boolean> }) => {
          commands.push(command);
        },
      }) as never,
    signalPluginAutoplayReady: () => {
      commands.push(['script-message', 'subminer-autoplay-ready']);
    },
    isSignalTargetReady: () => targetReady,
    schedule: (callback) => {
      scheduled.push(callback);
      return 1 as never;
    },
    logDebug: () => {},
  });

  gate.maybeSignalPluginAutoplayReady({ text: '字幕', tokens: null }, { forceWhilePaused: true });
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(commands, []);
  assert.equal(scheduled.length, 1);

  targetReady = true;
  scheduled.shift()?.();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(
    commands.filter((command) => command[0] === 'script-message'),
    [['script-message', 'subminer-autoplay-ready']],
  );
  assert.equal(
    commands.some(
      (command) => command[0] === 'set_property' && command[1] === 'pause' && command[2] === false,
    ),
    true,
  );
});

test('autoplay ready gate keeps deferred startup readiness retries active for cold starts', async () => {
  const commands: Array<Array<string | boolean>> = [];
  const scheduled: Array<() => void> = [];

  const gate = createAutoplayReadyGate({
    isAppOwnedFlowInFlight: () => false,
    getCurrentMediaPath: () => '/media/video.mkv',
    getCurrentVideoPath: () => null,
    getPlaybackPaused: () => true,
    getMpvClient: () =>
      ({
        connected: true,
        requestProperty: async () => true,
        send: ({ command }: { command: Array<string | boolean> }) => {
          commands.push(command);
        },
      }) as never,
    signalPluginAutoplayReady: () => {
      commands.push(['script-message', 'subminer-autoplay-ready']);
    },
    isSignalTargetReady: () => false,
    schedule: (callback) => {
      scheduled.push(callback);
      return 1 as never;
    },
    logDebug: () => {},
  });

  gate.maybeSignalPluginAutoplayReady(
    { text: '__warm__', tokens: null },
    { forceWhilePaused: true },
  );
  await new Promise((resolve) => setTimeout(resolve, 0));

  for (let attempt = 1; attempt <= 100; attempt += 1) {
    assert.equal(scheduled.length, 1, `missing deferred readiness retry ${attempt}`);
    scheduled.shift()?.();
    await new Promise((resolve) => setTimeout(resolve, 0));
  }

  assert.deepEqual(commands, []);
});

test('autoplay ready gate drops deferred readiness after media changes before flush', async () => {
  const commands: Array<Array<string | boolean>> = [];
  let targetReady = false;
  let currentMediaPath = '/media/video-1.mkv';

  const gate = createAutoplayReadyGate({
    isAppOwnedFlowInFlight: () => false,
    getCurrentMediaPath: () => currentMediaPath,
    getCurrentVideoPath: () => null,
    getPlaybackPaused: () => true,
    getMpvClient: () =>
      ({
        connected: true,
        requestProperty: async () => true,
        send: ({ command }: { command: Array<string | boolean> }) => {
          commands.push(command);
        },
      }) as never,
    signalPluginAutoplayReady: () => {
      commands.push(['script-message', 'subminer-autoplay-ready']);
    },
    isSignalTargetReady: () => targetReady,
    schedule: (callback) => {
      queueMicrotask(callback);
      return 1 as never;
    },
    logDebug: () => {},
  });

  gate.maybeSignalPluginAutoplayReady({ text: '字幕', tokens: null }, { forceWhilePaused: true });
  await new Promise((resolve) => setTimeout(resolve, 0));

  currentMediaPath = '/media/video-2.mkv';
  targetReady = true;
  gate.flushPendingAutoplayReadySignal();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(commands, []);
});

test('autoplay ready gate passes the pending subtitle signal to the readiness predicate', async () => {
  const commands: Array<Array<string | boolean>> = [];
  let targetReadyText: string | null = null;
  let observedText: string | null = null;
  let observedRequestedAtMs: number | null = null;
  let now = 1_000;

  const gate = createAutoplayReadyGate({
    isAppOwnedFlowInFlight: () => false,
    getCurrentMediaPath: () => '/media/video.mkv',
    getCurrentVideoPath: () => null,
    getPlaybackPaused: () => true,
    getMpvClient: () =>
      ({
        connected: true,
        requestProperty: async () => true,
        send: ({ command }: { command: Array<string | boolean> }) => {
          commands.push(command);
        },
      }) as never,
    signalPluginAutoplayReady: () => {
      commands.push(['script-message', 'subminer-autoplay-ready']);
    },
    isSignalTargetReady: ((signal: { payload: { text: string }; requestedAtMs: number }) => {
      observedText = signal.payload.text;
      observedRequestedAtMs = signal.requestedAtMs;
      return targetReadyText === signal.payload.text;
    }) as never,
    now: () => now,
    schedule: (callback) => {
      queueMicrotask(callback);
      return 1 as never;
    },
    logDebug: () => {},
  });

  gate.maybeSignalPluginAutoplayReady({ text: '字幕', tokens: null }, { forceWhilePaused: true });
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.equal(observedText, '字幕');
  assert.equal(observedRequestedAtMs, 1_000);
  assert.deepEqual(commands, []);

  now = 2_000;
  targetReadyText = '字幕';
  gate.flushPendingAutoplayReadySignal();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.equal(observedRequestedAtMs, 1_000);
  assert.deepEqual(
    commands.filter((command) => command[0] === 'script-message'),
    [['script-message', 'subminer-autoplay-ready']],
  );
});
