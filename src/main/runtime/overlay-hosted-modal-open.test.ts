import assert from 'node:assert/strict';
import test from 'node:test';
import { openOverlayHostedModal, retryOverlayModalOpen } from './overlay-hosted-modal-open';

test('retryOverlayModalOpen skips the first send when already aborted', async () => {
  const controller = new AbortController();
  controller.abort();
  const unexpectedCall = () => assert.fail('aborted open must not send or wait');
  assert.equal(
    await retryOverlayModalOpen(
      { waitForModalOpen: unexpectedCall, logWarn: unexpectedCall },
      {
        modal: 'media-timing-review',
        timeoutMs: 4_000,
        retryWarning: 'retry',
        sendOpen: unexpectedCall,
        signal: controller.signal,
      },
    ),
    false,
  );
});

for (const abortOnWait of [1, 2]) {
  test(`retryOverlayModalOpen rejects an acknowledgement aborted during wait ${abortOnWait}`, async () => {
    const controller = new AbortController();
    let waitCalls = 0;
    let sendCalls = 0;
    const opened = await retryOverlayModalOpen(
      {
        waitForModalOpen: async () => {
          waitCalls += 1;
          if (waitCalls === abortOnWait) {
            controller.abort();
            return true;
          }
          return false;
        },
        logWarn: () => {},
      },
      {
        modal: 'media-timing-review',
        timeoutMs: 4_000,
        retryWarning: 'retry',
        sendOpen: () => {
          sendCalls += 1;
          return true;
        },
        signal: controller.signal,
      },
    );
    assert.equal(opened, false);
    assert.equal(sendCalls, abortOnWait);
    assert.equal(waitCalls, abortOnWait);
  });
}

test('retryOverlayModalOpen still retries other modals without a signal', async () => {
  let sendCalls = 0;
  const opened = await retryOverlayModalOpen(
    { waitForModalOpen: async () => sendCalls === 2, logWarn: () => {} },
    {
      modal: 'runtime-options',
      timeoutMs: 1_500,
      retryWarning: 'retry',
      sendOpen: () => {
        sendCalls += 1;
        return true;
      },
    },
  );
  assert.equal(opened, true);
  assert.equal(sendCalls, 2);
});

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
