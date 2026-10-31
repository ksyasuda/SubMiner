import assert from 'node:assert/strict';
import test from 'node:test';

import { PollingRunner } from './polling';

test('polling runner records new cards and reports each ready after enrichment', async () => {
  const originalDateNow = Date.now;
  const recordedCards: number[] = [];
  const events: string[] = [];
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
      processNewCard: async (noteId) => {
        events.push(`process ${noteId}`);
      },
      recordCardsAdded: (count) => {
        recordedCards.push(count);
      },
      onCardsReady: (noteIds) => {
        events.push(`ready ${noteIds.join(',')}`);
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
    assert.deepEqual(events, ['process 12', 'ready 12', 'process 13', 'ready 13']);
  } finally {
    Date.now = originalDateNow;
  }
});

test('polling runner keeps going after one note fails to enrich', async () => {
  const originalDateNow = Date.now;
  const events: string[] = [];
  const warnings: unknown[][] = [];
  let tracked = new Set<number>();
  const responses = [[10], [10, 12, 13]];
  try {
    Date.now = () => 120_000;
    const runner = new PollingRunner({
      getDeck: () => undefined,
      getPollingRate: () => 250,
      findNotes: async () => responses.shift() ?? [],
      shouldAutoUpdateNewCards: () => true,
      processNewCard: async (noteId) => {
        events.push(`process ${noteId}`);
        if (noteId === 12) throw new Error('overlay callback failed');
      },
      onCardsReady: (noteIds) => {
        events.push(`ready ${noteIds.join(',')}`);
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
      logWarn: (...args) => {
        warnings.push(args);
      },
    });

    await runner.pollOnce();
    await runner.pollOnce();

    assert.deepEqual(events, ['process 12', 'ready 12', 'process 13', 'ready 13']);
    assert.deepEqual(warnings, [['Failed to auto-enrich note 12:', 'overlay callback failed']]);
  } finally {
    Date.now = originalDateNow;
  }
});
