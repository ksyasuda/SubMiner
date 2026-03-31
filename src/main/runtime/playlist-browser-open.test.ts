import assert from 'node:assert/strict';
import test from 'node:test';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { openPlaylistBrowser } from './playlist-browser-open';

test('playlist browser open bootstraps overlay runtime before dispatching the modal event', () => {
  const calls: string[] = [];

  const opened = openPlaylistBrowser({
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
      });
      return true;
    },
  });

  assert.equal(opened, true);
  assert.deepEqual(calls, ['prereqs', 'windows', `send:${IPC_CHANNELS.event.playlistBrowserOpen}`]);
});
