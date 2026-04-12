import assert from 'node:assert/strict';
import test from 'node:test';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { openPlaylistBrowser } from './playlist-browser-open';

test('playlist browser open bootstraps overlay runtime and sends modal event with preferModalWindow', async () => {
  const calls: string[] = [];

  const opened = await openPlaylistBrowser({
    ensureOverlayStartupPrereqs: () => {
      calls.push('prereqs');
    },
    ensureOverlayWindowsReadyForVisibilityActions: () => {
      calls.push('windows');
    },
    sendToActiveOverlayWindow: (channel, payload, runtimeOptions) => {
      calls.push(`send:${channel}`);
      assert.equal(payload, undefined);
      assert.deepEqual(runtimeOptions, {
        restoreOnModalClose: 'playlist-browser',
        preferModalWindow: true,
      });
      return true;
    },
    waitForModalOpen: async () => true,
    logWarn: () => {},
  });

  assert.equal(opened, true);
  assert.deepEqual(calls, ['prereqs', 'windows', `send:${IPC_CHANNELS.event.playlistBrowserOpen}`]);
});

test('playlist browser open retries after first attempt timeout', async () => {
  let attempt = 0;
  const opened = await openPlaylistBrowser({
    ensureOverlayStartupPrereqs: () => {},
    ensureOverlayWindowsReadyForVisibilityActions: () => {},
    sendToActiveOverlayWindow: () => true,
    waitForModalOpen: async () => {
      attempt += 1;
      return attempt >= 2;
    },
    logWarn: () => {},
  });

  assert.equal(opened, true);
  assert.equal(attempt, 2);
});
