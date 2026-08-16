import assert from 'node:assert/strict';
import test from 'node:test';
import {
  constrainMediaTimingSelection,
  createMediaTimingPreviewRequestGuard,
  formatMediaTimingTimestamp,
} from './media-timing-review';

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
