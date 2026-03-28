import assert from 'node:assert/strict';
import test from 'node:test';

import { PollingRunner } from './polling';

test('polling runner records newly added cards after initialization', async () => {
  const originalDateNow = Date.now;
  const recordedCards: number[] = [];
  let tracked = new Set<number>();
  const responses = [
    [10, 11],
    [10, 11, 12, 13],
  ];
  try {
    Date.now = () => 120_000;
    const runner = new PollingRunner({
      getDeck: () => 'Mining',
      getPollingRate: () => 250,
      findNotes: async () => responses.shift() ?? [],
      shouldAutoUpdateNewCards: () => true,
      processNewCard: async () => undefined,
      recordCardsAdded: (count) => {
        recordedCards.push(count);
      },
      isUpdateInProgress: () => false,
      setUpdateInProgress: () => undefined,
      getTrackedNoteIds: () => tracked,
      setTrackedNoteIds: (noteIds) => {
        tracked = noteIds;
      },
      showStatusNotification: () => undefined,
      logDebug: () => undefined,
      logInfo: () => undefined,
      logWarn: () => undefined,
    });

    await runner.pollOnce();
    await runner.pollOnce();

    assert.deepEqual(recordedCards, [2]);
  } finally {
    Date.now = originalDateNow;
  }
});
