import assert from 'node:assert/strict';
import test from 'node:test';

import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { openSessionHelpModal } from './session-help-open';

test('session help open tells the renderer whether commands can run', async () => {
  for (const playing of [true, false]) {
    const sent: Array<{ channel: string; payload: unknown }> = [];
    const opened = await openSessionHelpModal({
      ensureOverlayStartupPrereqs: () => {},
      ensureOverlayWindowsReadyForVisibilityActions: () => {},
      sendToActiveOverlayWindow: (channel, payload) => {
        sent.push({ channel, payload });
        return true;
      },
      waitForModalOpen: async () => true,
      logWarn: () => {},
      isMediaPlaybackActive: () => playing,
    });

    assert.equal(opened, true);
    assert.deepEqual(sent, [
      { channel: IPC_CHANNELS.event.sessionHelpOpen, payload: { commandsEnabled: playing } },
    ]);
  }
});
