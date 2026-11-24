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
      findNotes: async (query) => (query.includes('tag:') ? [] : (responses.shift() ?? [])),
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
      findNotes: async (query) => (query.includes('tag:') ? [] : (responses.shift() ?? [])),
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

function createStatsHarness(autoUpdate: boolean, sourceLookupFails = false) {
  let currentIds: number[] = [];
  let tracked = new Set<number>();
  const recorded: Array<{ count: number; noteIds: number[] }> = [];
  const processed: number[] = [];
  const ready: number[][] = [];
  const warnings: unknown[][] = [];
  const queries: string[] = [];
  const runner = new PollingRunner({
    getDeck: () => 'Mining',
    getPollingRate: () => 250,
    findNotes: async (query) => {
      queries.push(query);
      if (!query.includes('tag:')) return currentIds;
      if (sourceLookupFails) throw new Error('source lookup failed');
      return [2];
    },
    shouldAutoUpdateNewCards: () => autoUpdate,
    processNewCard: async (noteId) => {
      processed.push(noteId);
    },
    recordCardsAdded: (count, noteIds) => {
      recorded.push({ count, noteIds });
    },
    onCardsReady: (noteIds) => {
      ready.push(noteIds);
    },
    isUpdateInProgress: () => false,
    setUpdateInProgress: () => {},
    getTrackedNoteIds: () => tracked,
    setTrackedNoteIds: (noteIds) => {
      tracked = noteIds;
    },
    showStatusNotification: () => assert.fail('source lookup failures must not back off polling'),
    logDebug: () => {},
    logInfo: () => {},
    logWarn: (...args) => {
      warnings.push(args);
    },
  });
  return {
    recorded,
    processed,
    ready,
    warnings,
    queries,
    poll: async (noteIds: number[]) => {
      currentIds = noteIds;
      await runner.pollOnce();
    },
  };
}

for (const autoUpdate of [true, false]) {
  test(`stats cards count without triggering plugins when auto-update is ${autoUpdate}`, async () => {
    const h = createStatsHarness(autoUpdate);
    await h.poll([]);
    await h.poll([1, 2]);
    assert.deepEqual(h.recorded, [{ count: 2, noteIds: [1, 2] }]);
    assert.deepEqual(h.processed, autoUpdate ? [1, 2] : []);
    assert.deepEqual(h.ready, [[1]]);
    assert.ok(h.queries.includes('nid:1,2 tag:"SubMiner::Stats"'));
    assert.deepEqual(h.warnings, []);
  });
}

test('a failed source lookup skips plugin events without stopping counting or enrichment', async () => {
  const h = createStatsHarness(true, true);
  await h.poll([]);
  await h.poll([1, 2]);
  await h.poll([1, 2, 3]);
  assert.deepEqual(h.processed, [1, 2, 3]);
  assert.deepEqual(h.recorded, [
    { count: 2, noteIds: [1, 2] },
    { count: 1, noteIds: [3] },
  ]);
  assert.deepEqual(h.ready, []);
  assert.equal(h.warnings.length, 2);
});

/** A runner whose `findNotes` read stays pending until `finishRead`, sharing one update lock. */
function makeSlowReadRunner() {
  const state = { updating: false, processed: [] as number[], tracked: new Set<number>() };
  let finishRead: (noteIds: number[]) => void = () => undefined;
  const runner = new PollingRunner({
    getDeck: () => 'Mining',
    getPollingRate: () => 3000,
    findNotes: (query) =>
      query.includes('tag:')
        ? Promise.resolve([])
        : new Promise<number[]>((resolve) => {
            finishRead = resolve;
          }),
    shouldAutoUpdateNewCards: () => true,
    processNewCard: async (noteId) => {
      state.processed.push(noteId);
    },
    isUpdateInProgress: () => state.updating,
    setUpdateInProgress: (value) => {
      state.updating = value;
    },
    getTrackedNoteIds: () => state.tracked,
    setTrackedNoteIds: (noteIds) => {
      state.tracked = noteIds;
    },
    showStatusNotification: () => undefined,
    logDebug: () => undefined,
    logInfo: () => undefined,
    logWarn: () => undefined,
  });
  return {
    runner,
    state,
    finishRead: (noteIds: number[]) => finishRead(noteIds),
    /** The first poll only records the notes already in the deck. */
    initialize: async () => {
      const poll = runner.pollOnce();
      finishRead([10]);
      await poll;
    },
  };
}

test('a poll waiting on Anki leaves the update lock free for mining', async () => {
  const h = makeSlowReadRunner();
  await h.initialize();

  const poll = h.runner.pollOnce();
  assert.equal(h.state.updating, false, 'mining can start while the read is pending');
  h.finishRead([10, 11]);
  await poll;

  assert.deepEqual(h.state.processed, [11]);
  assert.equal(h.state.updating, false);
});

test('a poll leaves new notes alone when mining starts during its read', async () => {
  const h = makeSlowReadRunner();
  await h.initialize();

  const poll = h.runner.pollOnce();
  // Mining takes the lock and adds note 11 while the poll's read is in flight.
  h.state.updating = true;
  h.finishRead([10, 11]);
  await poll;

  assert.deepEqual(h.state.processed, []);
  assert.deepEqual([...h.state.tracked], [10]);
  assert.equal(h.state.updating, true, 'the poll must not release a lock it never took');
});
