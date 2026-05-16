import assert from 'node:assert/strict';
import test from 'node:test';
import { openSubsyncManualModal } from './subsync-open';
import type { SubsyncManualPayload } from '../../types';

const payload: SubsyncManualPayload = {
  sourceTracks: [{ id: 2, label: 'External #2 - eng' }],
};

test('subsync manual open prefers dedicated modal window on first attempt', async () => {
  const sends: Array<{
    channel: string;
    payload: SubsyncManualPayload;
    options: {
      restoreOnModalClose: 'subsync';
      preferModalWindow: boolean;
    };
  }> = [];

  const opened = await openSubsyncManualModal(
    {
      ensureOverlayStartupPrereqs: () => {},
      ensureOverlayWindowsReadyForVisibilityActions: () => {},
      sendToActiveOverlayWindow: (channel, nextPayload, options) => {
        sends.push({
          channel,
          payload: nextPayload as SubsyncManualPayload,
          options: options as {
            restoreOnModalClose: 'subsync';
            preferModalWindow: boolean;
          },
        });
        return true;
      },
      waitForModalOpen: async (modal, timeoutMs) => {
        assert.equal(modal, 'subsync');
        assert.equal(timeoutMs, 1500);
        return true;
      },
      logWarn: () => {
        throw new Error('should not warn on first-attempt success');
      },
    },
    payload,
  );

  assert.equal(opened, true);
  assert.deepEqual(sends, [
    {
      channel: 'subsync:open-manual',
      payload,
      options: {
        restoreOnModalClose: 'subsync',
        preferModalWindow: true,
      },
    },
  ]);
});

test('subsync manual open retries on the dedicated modal window after open timeout', async () => {
  const preferModalWindowValues: boolean[] = [];
  const warnings: string[] = [];
  let waitCalls = 0;

  const opened = await openSubsyncManualModal(
    {
      ensureOverlayStartupPrereqs: () => {},
      ensureOverlayWindowsReadyForVisibilityActions: () => {},
      sendToActiveOverlayWindow: (_channel, _payload, options) => {
        preferModalWindowValues.push(Boolean(options?.preferModalWindow));
        return true;
      },
      waitForModalOpen: async () => {
        waitCalls += 1;
        return waitCalls === 2;
      },
      logWarn: (message) => {
        warnings.push(message);
      },
    },
    payload,
  );

  assert.equal(opened, true);
  assert.deepEqual(preferModalWindowValues, [true, true]);
  assert.deepEqual(warnings, [
    'Subsync modal did not acknowledge modal open on first attempt; retrying dedicated modal window.',
  ]);
});

test('subsync manual open fails when the dedicated modal window cannot be targeted', async () => {
  let waitCalls = 0;

  const opened = await openSubsyncManualModal(
    {
      ensureOverlayStartupPrereqs: () => {},
      ensureOverlayWindowsReadyForVisibilityActions: () => {},
      sendToActiveOverlayWindow: () => false,
      waitForModalOpen: async () => {
        waitCalls += 1;
        return true;
      },
      logWarn: () => {},
    },
    payload,
  );

  assert.equal(opened, false);
  assert.equal(waitCalls, 0);
});
