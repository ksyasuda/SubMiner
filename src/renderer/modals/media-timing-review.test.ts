import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildMediaTimingLineSelection,
  buildMediaTimingWaveformPath,
  constrainMediaTimingSelection,
  createMediaTimingPreviewRequestGuard,
  formatMediaTimingTimestamp,
  mediaTimingTimeFromPointer,
  slideMediaTimingSelection,
} from './media-timing-review';

test('waveform path mirrors normalized peaks around its center line', () => {
  const path = buildMediaTimingWaveformPath([0, 0.5, 1]);

  assert.match(path, /^M 0\.00 50\.00 L 500\.00 28\.00 L 1000\.00 6\.00/);
  assert.match(path, /1000\.00 94\.00 L 500\.00 72\.00 L 0\.00 50\.00 Z$/);
  assert.equal(buildMediaTimingWaveformPath([1]), '');
});

test('formatMediaTimingTimestamp renders stable review readouts', () => {
  assert.equal(formatMediaTimingTimestamp(65.4321), '01:05.432');
  assert.equal(formatMediaTimingTimestamp(65.4321, false), '01:05');
  assert.equal(formatMediaTimingTimestamp(-1), '00:00.000');
  assert.equal(formatMediaTimingTimestamp(119.9999), '02:00.000');
  assert.equal(formatMediaTimingTimestamp(119.6, false), '02:00');
});

test('selection constraints preserve the handle that did not move', () => {
  assert.deepEqual(
    constrainMediaTimingSelection({
      nextStart: 3,
      nextEnd: 2,
      currentStart: 1,
      timelineStart: 0,
      timelineEnd: 10,
      mediaEnd: 10,
      maxMediaDuration: 30,
    }),
    { start: 1.9, end: 2 },
  );
  assert.deepEqual(
    constrainMediaTimingSelection({
      nextStart: 1,
      nextEnd: 0.5,
      currentStart: 1,
      timelineStart: 0,
      timelineEnd: 10,
      mediaEnd: 10,
      maxMediaDuration: 30,
    }),
    { start: 1, end: 1.1 },
  );
});

test('pointer positions map onto the visible timeline and clamp past its edges', () => {
  const track = { trackLeft: 100, trackWidth: 400, timelineStart: 10, timelineEnd: 20 };

  assert.equal(mediaTimingTimeFromPointer({ clientX: 100, ...track }), 10);
  assert.equal(mediaTimingTimeFromPointer({ clientX: 300, ...track }), 15);
  assert.equal(mediaTimingTimeFromPointer({ clientX: -40, ...track }), 10);
  assert.equal(mediaTimingTimeFromPointer({ clientX: 900, ...track }), 20);
  assert.equal(mediaTimingTimeFromPointer({ ...track, clientX: 300, trackWidth: 0 }), 10);
});

test('sliding keeps the clip length and stops at the timeline and media bounds', () => {
  const timeline = { span: 2, timelineStart: 4, timelineEnd: 12, mediaEnd: 10 };

  assert.deepEqual(slideMediaTimingSelection({ nextStart: 6, ...timeline }), { start: 6, end: 8 });
  assert.deepEqual(slideMediaTimingSelection({ nextStart: 1, ...timeline }), { start: 4, end: 6 });
  assert.deepEqual(slideMediaTimingSelection({ nextStart: 99, ...timeline }), {
    start: 8,
    end: 10,
  });
});

test('line selection combines adjacent lines around the mined one and tracks their range', () => {
  const base = {
    previousLines: [
      { text: '一行目', startTime: 0, endTime: 2 },
      { text: '二行目', startTime: 2.5, endTime: 4 },
    ],
    nextLines: [
      { text: '四行目', startTime: 7.5, endTime: 9 },
      { text: '五行目', startTime: 9.5, endTime: 11 },
    ],
    text: '採掘行',
    originalStartTime: 5,
    originalEndTime: 7,
  };

  const none = buildMediaTimingLineSelection({ ...base, previousCount: 0, nextCount: 0 });
  assert.deepEqual(none.lineTexts, ['採掘行']);
  assert.equal(none.currentLineIndex, 0);
  assert.equal(none.rangeStart, 5);
  assert.equal(none.rangeEnd, 7);

  const expanded = buildMediaTimingLineSelection({ ...base, previousCount: 1, nextCount: 2 });
  assert.deepEqual(expanded.lineTexts, ['二行目', '採掘行', '四行目', '五行目']);
  assert.equal(expanded.currentLineIndex, 1);
  assert.equal(expanded.sentence, '二行目 採掘行 四行目 五行目');
  assert.equal(expanded.rangeStart, 2.5);
  assert.equal(expanded.rangeEnd, 11);
});

test('preview request guard blocks overlap and invalidates stale responses', () => {
  const guard = createMediaTimingPreviewRequestGuard();
  const first = guard.begin();

  assert.equal(typeof first, 'number');
  assert.equal(guard.begin(), null);
  assert.equal(guard.isCurrent(first!), true);

  guard.invalidate();
  assert.equal(guard.isCurrent(first!), false);
  assert.equal(guard.isInFlight(), false);

  const second = guard.begin();
  assert.equal(typeof second, 'number');
  guard.finish(first!);
  assert.equal(guard.isCurrent(second!), true);
  guard.finish(second!);
  assert.equal(guard.isInFlight(), false);
});
