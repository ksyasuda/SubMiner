import assert from 'node:assert/strict';
import test from 'node:test';
import { EventType } from '../types/stats';
import {
  buildSessionChartEvents,
  extractSessionEventNoteInfo,
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

test('session marker pin helpers prefer pinned markers and toggle on repeat clicks', () => {
  assert.equal(resolveActiveSessionMarkerKey('card-1', 'seek-2'), 'seek-2');
  assert.equal(resolveActiveSessionMarkerKey('card-1', null), 'card-1');
  assert.equal(togglePinnedSessionMarkerKey(null, 'card-1'), 'card-1');
  assert.equal(togglePinnedSessionMarkerKey('card-1', 'card-1'), null);
  assert.equal(togglePinnedSessionMarkerKey('card-1', 'seek-2'), 'seek-2');
});
