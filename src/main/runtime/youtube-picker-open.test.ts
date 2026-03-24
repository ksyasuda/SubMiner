import assert from 'node:assert/strict';
import test from 'node:test';
import { openYoutubeTrackPicker } from './youtube-picker-open';
import type { YoutubePickerOpenPayload } from '../../types';

const payload: YoutubePickerOpenPayload = {
  sessionId: 'yt-1',
  url: 'https://example.com/watch?v=abc',
  tracks: [],
  defaultPrimaryTrackId: null,
  defaultSecondaryTrackId: null,
  hasTracks: false,
};

test('youtube picker open prefers dedicated modal window on first attempt', async () => {
  const sends: Array<{
    channel: string;
    payload: YoutubePickerOpenPayload;
    options: {
      restoreOnModalClose: 'youtube-track-picker';
      preferModalWindow: boolean;
    };
  }> = [];

  const opened = await openYoutubeTrackPicker(
    {
      sendToActiveOverlayWindow: (channel, nextPayload, options) => {
        sends.push({
          channel,
          payload: nextPayload as YoutubePickerOpenPayload,
          options: options as {
            restoreOnModalClose: 'youtube-track-picker';
            preferModalWindow: boolean;
          },
        });
        return true;
      },
      waitForModalOpen: async () => true,
      logWarn: () => {},
    },
    payload,
  );

  assert.equal(opened, true);
  assert.deepEqual(sends, [
    {
      channel: 'youtube:picker-open',
      payload,
      options: {
        restoreOnModalClose: 'youtube-track-picker',
        preferModalWindow: true,
      },
    },
  ]);
});

test('youtube picker open retries on the dedicated modal window after open timeout', async () => {
  const preferModalWindowValues: boolean[] = [];
  const warns: string[] = [];
  let waitCalls = 0;

  const opened = await openYoutubeTrackPicker(
    {
      sendToActiveOverlayWindow: (_channel, _payload, options) => {
        preferModalWindowValues.push(Boolean(options?.preferModalWindow));
        return true;
      },
      waitForModalOpen: async () => {
        waitCalls += 1;
        return waitCalls === 2;
      },
      logWarn: (message) => {
        warns.push(message);
      },
    },
    payload,
  );

  assert.equal(opened, true);
  assert.deepEqual(preferModalWindowValues, [true, true]);
  assert.equal(
    warns.includes(
      'YouTube subtitle picker did not acknowledge modal open on first attempt; retrying dedicated modal window.',
    ),
    true,
  );
});

test('youtube picker open fails when the dedicated modal window cannot be targeted', async () => {
  const opened = await openYoutubeTrackPicker(
    {
      sendToActiveOverlayWindow: () => false,
      waitForModalOpen: async () => true,
      logWarn: () => {},
    },
    payload,
  );

  assert.equal(opened, false);
});
