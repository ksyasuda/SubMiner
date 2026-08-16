import assert from 'node:assert/strict';
import test from 'node:test';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { openAnimeBrowserModal } from './anime-browser-open';

test('anime browser open uses a dedicated player-bounded modal window', async () => {
  const calls: string[] = [];

  const opened = await openAnimeBrowserModal({
    ensureOverlayStartupPrereqs: () => calls.push('prereqs'),
    ensureOverlayWindowsReadyForVisibilityActions: () => calls.push('windows'),
    sendToActiveOverlayWindow: (channel, payload, runtimeOptions) => {
      calls.push(`send:${channel}`);
      assert.equal(payload, undefined);
      assert.deepEqual(runtimeOptions, {
        restoreOnModalClose: 'anime-browser',
        preferModalWindow: true,
      });
      return true;
    },
    waitForModalOpen: async () => true,
    logWarn: () => {},
  });

  assert.equal(opened, true);
  assert.deepEqual(calls, ['prereqs', 'windows', `send:${IPC_CHANNELS.event.animeBrowserOpen}`]);
});

test('anime browser open retries on a fresh modal window after a missed acknowledgement', async () => {
  let attempts = 0;
  const opened = await openAnimeBrowserModal({
    ensureOverlayStartupPrereqs: () => {},
    ensureOverlayWindowsReadyForVisibilityActions: () => {},
    sendToActiveOverlayWindow: () => true,
    waitForModalOpen: async () => {
      attempts += 1;
      return attempts === 2;
    },
    logWarn: () => {},
  });

  assert.equal(opened, true);
  assert.equal(attempts, 2);
});
