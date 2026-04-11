import assert from 'node:assert/strict';
import test from 'node:test';
import { openRuntimeOptionsModal } from './runtime-options-open';

test('runtime options open prefers dedicated modal window on first attempt', async () => {
  const calls: string[] = [];

  const opened = await openRuntimeOptionsModal({
    ensureOverlayStartupPrereqs: () => {
      calls.push('ensure-startup');
    },
    ensureOverlayWindowsReadyForVisibilityActions: () => {
      calls.push('ensure-windows');
    },
    sendToActiveOverlayWindow: (channel, payload, options) => {
      calls.push(`send:${channel}`);
      assert.equal(payload, undefined);
      assert.deepEqual(options, {
        restoreOnModalClose: 'runtime-options',
        preferModalWindow: true,
      });
      return true;
    },
    waitForModalOpen: async () => true,
    logWarn: () => {
      throw new Error('should not warn on first-attempt success');
    },
  });

  assert.equal(opened, true);
  assert.deepEqual(calls, ['ensure-startup', 'ensure-windows', 'send:runtime-options:open']);
});

test('runtime options open retries after an open timeout', async () => {
  const calls: string[] = [];
  const warnings: string[] = [];
  let waitCalls = 0;

  const opened = await openRuntimeOptionsModal({
    ensureOverlayStartupPrereqs: () => {
      calls.push('ensure-startup');
    },
    ensureOverlayWindowsReadyForVisibilityActions: () => {
      calls.push('ensure-windows');
    },
    sendToActiveOverlayWindow: (channel, payload, options) => {
      calls.push(`send:${channel}`);
      assert.equal(payload, undefined);
      assert.deepEqual(options, {
        restoreOnModalClose: 'runtime-options',
        preferModalWindow: true,
      });
      return true;
    },
    waitForModalOpen: async () => {
      waitCalls += 1;
      return waitCalls === 2;
    },
    logWarn: (message) => {
      warnings.push(message);
    },
  });

  assert.equal(opened, true);
  assert.deepEqual(calls, [
    'ensure-startup',
    'ensure-windows',
    'send:runtime-options:open',
    'ensure-startup',
    'ensure-windows',
    'send:runtime-options:open',
  ]);
  assert.deepEqual(warnings, [
    'Runtime options modal did not acknowledge modal open on first attempt; retrying dedicated modal window.',
  ]);
});

test('runtime options open fails when no overlay window can be targeted', async () => {
  const opened = await openRuntimeOptionsModal({
    ensureOverlayStartupPrereqs: () => {},
    ensureOverlayWindowsReadyForVisibilityActions: () => {},
    sendToActiveOverlayWindow: () => false,
    waitForModalOpen: async () => true,
    logWarn: () => {},
  });

  assert.equal(opened, false);
});
