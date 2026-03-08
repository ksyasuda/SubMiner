import test from 'node:test';
import assert from 'node:assert/strict';

import { runLoggedYoutubePhase } from './progress';

test('runLoggedYoutubePhase logs start and finish with elapsed time', async () => {
  const entries: Array<{ level: 'info' | 'warn'; message: string }> = [];
  let nowMs = 1_000;

  const result = await runLoggedYoutubePhase(
    {
      startMessage: 'Starting subtitle probe',
      finishMessage: 'Finished subtitle probe',
      log: (level, message) => entries.push({ level, message }),
      now: () => nowMs,
    },
    async () => {
      nowMs = 2_500;
      return 'ok';
    },
  );

  assert.equal(result, 'ok');
  assert.deepEqual(entries, [
    { level: 'info', message: 'Starting subtitle probe' },
    { level: 'info', message: 'Finished subtitle probe (1.5s)' },
  ]);
});

test('runLoggedYoutubePhase logs failure with elapsed time and rethrows', async () => {
  const entries: Array<{ level: 'info' | 'warn'; message: string }> = [];
  let nowMs = 5_000;

  await assert.rejects(
    runLoggedYoutubePhase(
      {
        startMessage: 'Starting whisper primary',
        finishMessage: 'Finished whisper primary',
        failureMessage: 'Failed whisper primary',
        log: (level, message) => entries.push({ level, message }),
        now: () => nowMs,
      },
      async () => {
        nowMs = 8_200;
        throw new Error('boom');
      },
    ),
    /boom/,
  );

  assert.deepEqual(entries, [
    { level: 'info', message: 'Starting whisper primary' },
    { level: 'warn', message: 'Failed whisper primary after 3.2s: boom' },
  ]);
});
