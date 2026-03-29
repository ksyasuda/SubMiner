import assert from 'node:assert/strict';
import test from 'node:test';
import {
  DEFAULT_AUTOPLAY_RELEASE_RETRY_DELAY_MS,
  resolveAutoplayReadyMaxReleaseAttempts,
  STARTUP_AUTOPLAY_RELEASE_TIMEOUT_MS,
} from './startup-autoplay-release-policy';

test('autoplay release keeps the short retry budget for normal playback signals', () => {
  assert.equal(resolveAutoplayReadyMaxReleaseAttempts(), 3);
  assert.equal(resolveAutoplayReadyMaxReleaseAttempts({ forceWhilePaused: false }), 3);
});

test('autoplay release uses the full startup timeout window while paused', () => {
  assert.equal(
    resolveAutoplayReadyMaxReleaseAttempts({ forceWhilePaused: true }),
    Math.ceil(STARTUP_AUTOPLAY_RELEASE_TIMEOUT_MS / DEFAULT_AUTOPLAY_RELEASE_RETRY_DELAY_MS),
  );
});

test('autoplay release rounds up custom paused retry budgets to cover the timeout window', () => {
  assert.equal(
    resolveAutoplayReadyMaxReleaseAttempts({
      forceWhilePaused: true,
      retryDelayMs: 300,
      startupTimeoutMs: 1_000,
    }),
    4,
  );
});
