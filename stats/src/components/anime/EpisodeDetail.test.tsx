import assert from 'node:assert/strict';
import test from 'node:test';
import { filterCardEvents } from './EpisodeDetail';
import type { EpisodeCardEvent } from '../../types/stats';

function makeEvent(over: Partial<EpisodeCardEvent> & { eventId: number }): EpisodeCardEvent {
  return {
    sessionId: 1,
    tsMs: 0,
    cardsDelta: 1,
    noteIds: [],
    ...over,
  };
}

test('filterCardEvents: before load, returns all events unchanged', () => {
  const ev1 = makeEvent({ eventId: 1, noteIds: [101] });
  const ev2 = makeEvent({ eventId: 2, noteIds: [102] });
  const noteInfos = new Map(); // empty — simulates pre-load state
  const result = filterCardEvents([ev1, ev2], noteInfos, /* noteInfosLoaded */ false);
  assert.equal(result.length, 2, 'should return both events before load');
  assert.deepEqual(result[0]?.noteIds, [101]);
  assert.deepEqual(result[1]?.noteIds, [102]);
});

test('filterCardEvents: after load, drops noteIds not in noteInfos', () => {
  const ev1 = makeEvent({ eventId: 1, noteIds: [101] }); // survives
  const ev2 = makeEvent({ eventId: 2, noteIds: [102] }); // deleted from Anki
  const noteInfos = new Map([[101, { noteId: 101, expression: '食べる' }]]);
  const result = filterCardEvents([ev1, ev2], noteInfos, /* noteInfosLoaded */ true);
  assert.equal(result.length, 1, 'should drop event whose noteId was deleted from Anki');
  assert.equal(result[0]?.eventId, 1);
  assert.deepEqual(result[0]?.noteIds, [101]);
});

test('filterCardEvents: after load, legacy rollup events (empty noteIds, positive cardsDelta) are kept', () => {
  const rollup = makeEvent({ eventId: 3, noteIds: [], cardsDelta: 5 });
  const noteInfos = new Map<number, { noteId: number; expression: string }>();
  const result = filterCardEvents([rollup], noteInfos, true);
  assert.equal(result.length, 1, 'legacy rollup event should survive filtering');
  assert.equal(result[0]?.cardsDelta, 5);
});

test('filterCardEvents: after load, event with multiple noteIds keeps surviving ones', () => {
  const ev = makeEvent({ eventId: 4, noteIds: [201, 202, 203] });
  const noteInfos = new Map([
    [201, { noteId: 201, expression: 'A' }],
    [203, { noteId: 203, expression: 'C' }],
  ]);
  const result = filterCardEvents([ev], noteInfos, true);
  assert.equal(result.length, 1, 'event with surviving noteIds should be kept');
  assert.deepEqual(result[0]?.noteIds, [201, 203], 'only surviving noteIds should remain');
});

test('filterCardEvents: after load, event where all noteIds deleted is dropped', () => {
  const ev = makeEvent({ eventId: 5, noteIds: [301, 302] });
  const noteInfos = new Map<number, { noteId: number; expression: string }>();
  const result = filterCardEvents([ev], noteInfos, true);
  assert.equal(result.length, 0, 'event with all noteIds deleted should be dropped');
});
