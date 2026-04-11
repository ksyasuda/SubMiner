import assert from 'node:assert/strict';
import test from 'node:test';
import { openOverlayHostedModal } from './overlay-hosted-modal-open';

test('openOverlayHostedModal ensures overlay readiness before sending the open event', () => {
  const calls: string[] = [];

  const opened = openOverlayHostedModal(
    {
      ensureOverlayStartupPrereqs: () => {
        calls.push('ensureOverlayStartupPrereqs');
      },
      ensureOverlayWindowsReadyForVisibilityActions: () => {
        calls.push('ensureOverlayWindowsReadyForVisibilityActions');
      },
      sendToActiveOverlayWindow: (channel, payload, runtimeOptions) => {
        calls.push(`send:${channel}`);
        assert.equal(payload, undefined);
        assert.deepEqual(runtimeOptions, {
          restoreOnModalClose: 'runtime-options',
          preferModalWindow: undefined,
        });
        return true;
      },
    },
    {
      channel: 'runtime-options:open',
      modal: 'runtime-options',
    },
  );

  assert.equal(opened, true);
  assert.deepEqual(calls, [
    'ensureOverlayStartupPrereqs',
    'ensureOverlayWindowsReadyForVisibilityActions',
    'send:runtime-options:open',
  ]);
});

test('openOverlayHostedModal forwards payload and modal-window preference', () => {
  const payload = { sessionId: 'yt-1' };

  const opened = openOverlayHostedModal(
    {
      ensureOverlayStartupPrereqs: () => {},
      ensureOverlayWindowsReadyForVisibilityActions: () => {},
      sendToActiveOverlayWindow: (channel, forwardedPayload, runtimeOptions) => {
        assert.equal(channel, 'youtube:picker-open');
        assert.deepEqual(forwardedPayload, payload);
        assert.deepEqual(runtimeOptions, {
          restoreOnModalClose: 'youtube-track-picker',
          preferModalWindow: true,
        });
        return false;
      },
    },
    {
      channel: 'youtube:picker-open',
      modal: 'youtube-track-picker',
      payload,
      preferModalWindow: true,
    },
  );

  assert.equal(opened, false);
});
