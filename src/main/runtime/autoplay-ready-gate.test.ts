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
  assert.equal(scheduled.length > 0, true);
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
