import assert from 'node:assert/strict';
import test from 'node:test';
import { restoreMacOSMpvFocusAfterModalClose } from './macos-modal-focus-handoff';

test('restoreMacOSMpvFocusAfterModalClose focuses mpv, refreshes tracker, then updates visibility on macOS', async () => {
  const calls: string[] = [];

  await restoreMacOSMpvFocusAfterModalClose({
    platform: 'darwin',
    focusMpv: async () => {
      calls.push('focus');
    },
    getWindowTracker: () => ({
      refreshNow: async () => {
        calls.push('refresh');
      },
    }),
    updateVisibleOverlayVisibility: () => {
      calls.push('visibility');
    },
    warn: () => {},
  });

  assert.deepEqual(calls, ['focus', 'refresh', 'visibility']);
});

test('restoreMacOSMpvFocusAfterModalClose skips non-macOS platforms', async () => {
  const calls: string[] = [];

  await restoreMacOSMpvFocusAfterModalClose({
    platform: 'linux',
    focusMpv: async () => {
      calls.push('focus');
    },
    getWindowTracker: () => null,
    updateVisibleOverlayVisibility: () => {
      calls.push('visibility');
    },
    warn: () => {},
  });

  assert.deepEqual(calls, []);
});

test('restoreMacOSMpvFocusAfterModalClose still updates visibility when tracker refresh fails', async () => {
  const calls: string[] = [];

  await restoreMacOSMpvFocusAfterModalClose({
    platform: 'darwin',
    focusMpv: async () => {
      calls.push('focus');
    },
    getWindowTracker: () => ({
      refreshNow: async () => {
        calls.push('refresh');
        throw new Error('refresh failed');
      },
    }),
    updateVisibleOverlayVisibility: () => {
      calls.push('visibility');
    },
    warn: (message) => {
      calls.push(`warn:${message}`);
    },
  });

  assert.deepEqual(calls, [
    'focus',
    'refresh',
    'warn:Failed to refresh macOS mpv focus after modal close',
    'visibility',
  ]);
});
