import assert from 'node:assert/strict';
import test from 'node:test';
import { createAutoplayReadyGate } from './autoplay-ready-gate';

test('autoplay ready gate suppresses duplicate media signals unless forced while paused', async () => {
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
  gate.maybeSignalPluginAutoplayReady({ text: '字幕', tokens: null }, { forceWhilePaused: true });

  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(commands.slice(0, 3), [
    ['script-message', 'subminer-autoplay-ready'],
    ['script-message', 'subminer-autoplay-ready'],
    ['script-message', 'subminer-autoplay-ready'],
  ]);
  assert.ok(commands.some((command) => command[0] === 'set_property' && command[1] === 'pause'));
  assert.equal(scheduled.length > 0, true);
});
