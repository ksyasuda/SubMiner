import test from 'node:test';
import assert from 'node:assert/strict';
import {
  beginUpdateProgress,
  createUiFeedbackState,
  showProgressTick,
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
