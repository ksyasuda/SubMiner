import assert from 'node:assert/strict';
import test from 'node:test';
import { openSubtitleGenerationModal } from './subtitle-generation-open';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';

test('subtitle generation opens in the dedicated modal window with normal close restoration', async () => {
  const calls: string[] = [];
  const opened = await openSubtitleGenerationModal({
    ensureOverlayStartupPrereqs: () => {
      calls.push('startup');
    },
    ensureOverlayWindowsReadyForVisibilityActions: () => {
      calls.push('windows');
    },
    sendToActiveOverlayWindow: (channel, payload, options) => {
      assert.deepEqual(calls, ['startup', 'windows']);
      assert.equal(channel, IPC_CHANNELS.event.subtitleGenerationOpen);
      assert.equal(payload, undefined);
      assert.deepEqual(options, {
        restoreOnModalClose: 'subtitle-generation',
        preferModalWindow: true,
      });
      calls.push('open');
      return true;
    },
    waitForModalOpen: async (modal) => {
      assert.equal(modal, 'subtitle-generation');
      return true;
    },
    logWarn: () => {
      assert.fail('opening should not require a retry');
    },
  });
  assert.equal(opened, true);
  assert.deepEqual(calls, ['startup', 'windows', 'open']);
});
