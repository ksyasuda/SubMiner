import assert from 'node:assert/strict';
import test from 'node:test';
import {
  beginUpdateProgress,
  createUiFeedbackState,
  showProgressTick,
  showStatusNotification,
  showUpdateResult,
} from './ui-feedback';

test('showUpdateResult stops spinner before success notification and suppresses stale ticks', () => {
  const state = createUiFeedbackState();
  const osdMessages: string[] = [];

  beginUpdateProgress(state, 'Creating sentence card', () => {
    showProgressTick(state, (text) => {
      osdMessages.push(text);
    });
  });

  showUpdateResult(
    state,
    {
      clearProgressTimer: (timer) => {
        clearInterval(timer);
      },
      showOsdNotification: (text) => {
        osdMessages.push(text);
      },
    },
    { success: true, message: 'Updated card: taberu' },
  );

  showProgressTick(state, (text) => {
    osdMessages.push(text);
  });

  assert.deepEqual(osdMessages, ['Creating sentence card |', '✓ Updated card: taberu']);
});

test('showUpdateResult renders failed updates with an x marker', () => {
  const state = createUiFeedbackState();
  const osdMessages: string[] = [];

  beginUpdateProgress(state, 'Creating sentence card', () => {
    showProgressTick(state, (text) => {
      osdMessages.push(text);
    });
  });

  showUpdateResult(
    state,
    {
      clearProgressTimer: (timer) => {
        clearInterval(timer);
      },
      showOsdNotification: (text) => {
        osdMessages.push(text);
      },
    },
    { success: false, message: 'Sentence card failed: deck missing' },
  );

  assert.deepEqual(osdMessages, [
    'Creating sentence card |',
    'x Sentence card failed: deck missing',
  ]);
});

test('showStatusNotification falls back to system when overlay delivery is unavailable', () => {
  const calls: string[] = [];

  showStatusNotification('Waiting for card update', {
    getNotificationType: () => 'overlay',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showSystemNotification: (title, options) => {
      calls.push(`system:${title}:${options.body}`);
    },
  });

  assert.deepEqual(calls, ['system:SubMiner:Waiting for card update']);
});

test('showStatusNotification defaults to mpv osd when notification type is unset', () => {
  const calls: string[] = [];

  showStatusNotification('Card updated', {
    getNotificationType: () => undefined,
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showOverlayNotification: (payload) => {
      calls.push(`overlay:${payload.body}`);
    },
    showSystemNotification: (title, options) => {
      calls.push(`system:${title}:${options.body}`);
    },
  });

  assert.deepEqual(calls, ['osd:Card updated']);
});

test('showStatusNotification does not duplicate system notifications for both', () => {
  const calls: string[] = [];

  showStatusNotification('Card updated', {
    getNotificationType: () => 'both',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showOverlayNotification: (payload) => {
      calls.push(`overlay:${payload.body}`);
    },
    showSystemNotification: (title, options) => {
      calls.push(`system:${title}:${options.body}`);
    },
  });

  assert.deepEqual(calls, ['overlay:Card updated', 'system:SubMiner:Card updated']);
});
