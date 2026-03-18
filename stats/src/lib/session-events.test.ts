import assert from 'node:assert/strict';
import test from 'node:test';
import { EventType } from '../types/stats';
import {
  buildSessionChartEvents,
  collectPendingSessionEventNoteIds,
  extractSessionEventNoteInfo,
  getSessionEventCardRequest,
  mergeSessionEventNoteInfos,
  projectSessionMarkerLeftPx,
  resolveActiveSessionMarkerKey,
  togglePinnedSessionMarkerKey,
} from './session-events';

test('buildSessionChartEvents produces typed hover markers with parsed payload metadata', () => {
  const chartEvents = buildSessionChartEvents([
    { eventType: EventType.PAUSE_START, tsMs: 2_000, payload: null },
    {
      eventType: EventType.SEEK_FORWARD,
      tsMs: 3_000,
      payload: '{"fromMs":1000,"toMs":5500}',
    },
    { eventType: EventType.PAUSE_END, tsMs: 5_000, payload: null },
    {
      eventType: EventType.CARD_MINED,
      tsMs: 6_000,
      payload: '{"cardsMined":2,"noteIds":[11,22]}',
    },
    { eventType: EventType.YOMITAN_LOOKUP, tsMs: 7_000, payload: null },
  ]);

  assert.deepEqual(
    chartEvents.markers.map((marker) => marker.kind),
    ['seek', 'pause', 'card'],
  );

  const seekMarker = chartEvents.markers[0]!;
  assert.equal(seekMarker.kind, 'seek');
  assert.equal(seekMarker.direction, 'forward');
  assert.equal(seekMarker.fromMs, 1_000);
  assert.equal(seekMarker.toMs, 5_500);

  const pauseMarker = chartEvents.markers[1]!;
  assert.equal(pauseMarker.kind, 'pause');
  assert.equal(pauseMarker.startMs, 2_000);
  assert.equal(pauseMarker.endMs, 5_000);
  assert.equal(pauseMarker.durationMs, 3_000);
  assert.equal(pauseMarker.anchorTsMs, 3_500);

  const cardMarker = chartEvents.markers[2]!;
  assert.equal(cardMarker.kind, 'card');
  assert.deepEqual(cardMarker.noteIds, [11, 22]);
  assert.equal(cardMarker.cardsDelta, 2);

  assert.deepEqual(
    chartEvents.yomitanLookupEvents.map((event) => event.tsMs),
    [7_000],
  );
});

test('projectSessionMarkerLeftPx respects chart plot offsets instead of full-width percentages', () => {
  assert.equal(
    projectSessionMarkerLeftPx({
      anchorTsMs: 1_000,
      tsMin: 1_000,
      tsMax: 11_000,
      plotLeftPx: 5,
      plotWidthPx: 958,
    }),
    5,
  );

  assert.equal(
    projectSessionMarkerLeftPx({
      anchorTsMs: 6_000,
      tsMin: 1_000,
      tsMax: 11_000,
      plotLeftPx: 5,
      plotWidthPx: 958,
    }),
    484,
  );

  assert.equal(
    projectSessionMarkerLeftPx({
      anchorTsMs: 11_000,
      tsMin: 1_000,
      tsMax: 11_000,
      plotLeftPx: 5,
      plotWidthPx: 958,
    }),
    963,
  );
});

test('extractSessionEventNoteInfo prefers expression-like fields and strips html', () => {
  const info = extractSessionEventNoteInfo({
    noteId: 91,
    fields: {
      Sentence: { value: '<div>この呪いの剣は危険だ</div>' },
      Vocabulary: { value: '<span>呪いの剣</span>' },
      Meaning: { value: '<div>cursed sword</div>' },
    },
  });

  assert.deepEqual(info, {
    noteId: 91,
    expression: '呪いの剣',
    context: 'この呪いの剣は危険だ',
    meaning: 'cursed sword',
  });
});

test('extractSessionEventNoteInfo prefers explicit preview payload over field-name guessing', () => {
  const info = extractSessionEventNoteInfo({
    noteId: 92,
    preview: {
      word: '連れる',
      sentence: 'このまま 連れてって',
      translation: 'to take along',
    },
    fields: {
      UnexpectedWordField: { value: 'should not win' },
      UnexpectedSentenceField: { value: 'should not win either' },
    },
  });

  assert.deepEqual(info, {
    noteId: 92,
    expression: '連れる',
    context: 'このまま 連れてって',
    meaning: 'to take along',
  });
});

test('extractSessionEventNoteInfo ignores malformed notes without a numeric note id', () => {
  assert.equal(
    extractSessionEventNoteInfo({
      noteId: Number.NaN,
      fields: {
        Vocabulary: { value: '呪い' },
      },
    }),
    null,
  );
});

test('mergeSessionEventNoteInfos keys previews by both requested and returned note ids', () => {
  const noteInfos = mergeSessionEventNoteInfos([111], [
    {
      noteId: 222,
      fields: {
        Expression: { value: '呪い' },
        Sentence: { value: 'この剣は呪いだ' },
      },
    },
  ]);

  assert.deepEqual(noteInfos.get(111), {
    noteId: 222,
    expression: '呪い',
    context: 'この剣は呪いだ',
    meaning: null,
  });
  assert.deepEqual(noteInfos.get(222), {
    noteId: 222,
    expression: '呪い',
    context: 'この剣は呪いだ',
    meaning: null,
  });
});

test('collectPendingSessionEventNoteIds supports strict-mode cleanup and refetch', () => {
  const noteInfos = new Map();
  const pendingNoteIds = new Set<number>();

  assert.deepEqual(collectPendingSessionEventNoteIds([177], noteInfos, pendingNoteIds), [177]);

  pendingNoteIds.add(177);
  assert.deepEqual(collectPendingSessionEventNoteIds([177], noteInfos, pendingNoteIds), []);

  pendingNoteIds.delete(177);
  assert.deepEqual(collectPendingSessionEventNoteIds([177], noteInfos, pendingNoteIds), [177]);

  noteInfos.set(177, {
    noteId: 177,
    expression: '対抗',
    context: 'ダクネス 無理して 対抗 するな',
    meaning: null,
  });
  assert.deepEqual(collectPendingSessionEventNoteIds([177], noteInfos, pendingNoteIds), []);
});

test('getSessionEventCardRequest stays stable across rebuilt marker objects', () => {
  const events = [
    {
      eventType: EventType.CARD_MINED,
      tsMs: 6_000,
      payload: '{"cardsMined":1,"noteIds":[1773808840964]}',
    },
  ];

  const firstMarker = buildSessionChartEvents(events).markers[0]!;
  const secondMarker = buildSessionChartEvents(events).markers[0]!;

  assert.notEqual(firstMarker, secondMarker);
  assert.deepEqual(getSessionEventCardRequest(firstMarker), {
    noteIds: [1773808840964],
    requestKey: 'card-6000:1773808840964',
  });
  assert.deepEqual(getSessionEventCardRequest(secondMarker), {
    noteIds: [1773808840964],
    requestKey: 'card-6000:1773808840964',
  });
});

test('session marker pin helpers prefer pinned markers and toggle on repeat clicks', () => {
  assert.equal(resolveActiveSessionMarkerKey('card-1', 'seek-2'), 'seek-2');
  assert.equal(resolveActiveSessionMarkerKey('card-1', null), 'card-1');
  assert.equal(togglePinnedSessionMarkerKey(null, 'card-1'), 'card-1');
  assert.equal(togglePinnedSessionMarkerKey('card-1', 'card-1'), null);
  assert.equal(togglePinnedSessionMarkerKey('card-1', 'seek-2'), 'seek-2');
});
