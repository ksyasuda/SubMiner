import test from 'node:test';
import assert from 'node:assert/strict';
import { WindowsWindowTracker } from './windows-tracker';
import type { MpvPollResult } from './win32';

function mpvVisible(
  overrides: Partial<MpvPollResult & { x?: number; y?: number; width?: number; height?: number; focused?: boolean }> = {},
): MpvPollResult {
  return {
    matches: [
      {
        hwnd: 12345,
        bounds: {
          x: overrides.x ?? 0,
          y: overrides.y ?? 0,
          width: overrides.width ?? 1280,
          height: overrides.height ?? 720,
        },
        area: (overrides.width ?? 1280) * (overrides.height ?? 720),
        isForeground: overrides.focused ?? true,
      },
    ],
    focusState: overrides.focused ?? true,
    windowState: 'visible',
  };
}

const mpvNotFound: MpvPollResult = {
  matches: [],
  focusState: false,
  windowState: 'not-found',
};

const mpvMinimized: MpvPollResult = {
  matches: [],
  focusState: false,
  windowState: 'minimized',
};

test('WindowsWindowTracker skips overlapping polls while poll is in flight', () => {
  let pollCalls = 0;
  let tracker: WindowsWindowTracker;
  tracker = new WindowsWindowTracker(undefined, {
    pollMpvWindows: () => {
      pollCalls += 1;
      if (pollCalls === 1) {
        (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
      }
      return mpvVisible();
    },
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(pollCalls, 1);
});

test('WindowsWindowTracker updates geometry from poll output', () => {
  const tracker = new WindowsWindowTracker(undefined, {
    pollMpvWindows: () => mpvVisible({ x: 10, y: 20, width: 1280, height: 720 }),
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();

  assert.deepEqual(tracker.getGeometry(), {
    x: 10,
    y: 20,
    width: 1280,
    height: 720,
  });
  assert.equal(tracker.isTargetWindowFocused(), true);
});

test('WindowsWindowTracker clears geometry for poll misses', () => {
  const tracker = new WindowsWindowTracker(undefined, {
    pollMpvWindows: () => mpvNotFound,
    trackingLossGraceMs: 0,
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(tracker.getGeometry(), null);
  assert.equal(tracker.isTargetWindowFocused(), false);
});

test('WindowsWindowTracker keeps the last geometry through a single poll miss', () => {
  let callIndex = 0;
  const outputs = [
    mpvVisible({ x: 10, y: 20, width: 1280, height: 720 }),
    mpvNotFound,
    mpvVisible({ x: 10, y: 20, width: 1280, height: 720 }),
  ];

  const tracker = new WindowsWindowTracker(undefined, {
    pollMpvWindows: () => outputs[callIndex++] ?? outputs.at(-1)!,
    trackingLossGraceMs: 0,
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.deepEqual(tracker.getGeometry(), { x: 10, y: 20, width: 1280, height: 720 });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.deepEqual(tracker.getGeometry(), { x: 10, y: 20, width: 1280, height: 720 });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.deepEqual(tracker.getGeometry(), { x: 10, y: 20, width: 1280, height: 720 });
});

test('WindowsWindowTracker drops tracking after grace window expires', () => {
  let callIndex = 0;
  let now = 1_000;
  const outputs = [
    mpvVisible({ x: 10, y: 20, width: 1280, height: 720 }),
    mpvNotFound,
    mpvNotFound,
    mpvNotFound,
    mpvNotFound,
  ];

  const tracker = new WindowsWindowTracker(undefined, {
    pollMpvWindows: () => outputs[callIndex++] ?? outputs.at(-1)!,
    now: () => now,
    trackingLossGraceMs: 500,
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(tracker.isTracking(), true);

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(tracker.isTracking(), true);

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(tracker.isTracking(), true);

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(tracker.isTracking(), true);

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(tracker.isTracking(), false);
  assert.equal(tracker.getGeometry(), null);
});

test('WindowsWindowTracker keeps tracking through repeated poll misses inside grace window', () => {
  let callIndex = 0;
  let now = 1_000;
  const outputs = [
    mpvVisible({ x: 10, y: 20, width: 1280, height: 720 }),
    mpvNotFound,
    mpvNotFound,
    mpvNotFound,
    mpvVisible({ x: 10, y: 20, width: 1280, height: 720 }),
  ];

  const tracker = new WindowsWindowTracker(undefined, {
    pollMpvWindows: () => outputs[callIndex++] ?? outputs.at(-1)!,
    now: () => now,
    trackingLossGraceMs: 1_500,
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(tracker.isTracking(), true);

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(tracker.isTracking(), true);

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(tracker.isTracking(), true);

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(tracker.isTracking(), true);
  assert.deepEqual(tracker.getGeometry(), { x: 10, y: 20, width: 1280, height: 720 });

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(tracker.isTracking(), true);
  assert.deepEqual(tracker.getGeometry(), { x: 10, y: 20, width: 1280, height: 720 });
});

test('WindowsWindowTracker keeps tracking through a transient minimized report inside minimized grace window', () => {
  let callIndex = 0;
  let now = 1_000;
  const outputs: MpvPollResult[] = [
    mpvVisible({ x: 10, y: 20, width: 1280, height: 720 }),
    mpvMinimized,
    mpvVisible({ x: 10, y: 20, width: 1280, height: 720 }),
  ];

  const tracker = new WindowsWindowTracker(undefined, {
    pollMpvWindows: () => outputs[callIndex++] ?? outputs.at(-1)!,
    now: () => now,
    minimizedTrackingLossGraceMs: 200,
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(tracker.isTracking(), true);

  now += 100;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(tracker.isTracking(), true);
  assert.deepEqual(tracker.getGeometry(), { x: 10, y: 20, width: 1280, height: 720 });

  now += 100;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(tracker.isTracking(), true);
  assert.deepEqual(tracker.getGeometry(), { x: 10, y: 20, width: 1280, height: 720 });
});

test('WindowsWindowTracker keeps tracking through repeated transient minimized reports inside minimized grace window', () => {
  let callIndex = 0;
  let now = 1_000;
  const outputs: MpvPollResult[] = [
    mpvVisible({ x: 10, y: 20, width: 1280, height: 720 }),
    mpvMinimized,
    mpvMinimized,
    mpvVisible({ x: 10, y: 20, width: 1280, height: 720 }),
  ];

  const tracker = new WindowsWindowTracker(undefined, {
    pollMpvWindows: () => outputs[callIndex++] ?? outputs.at(-1)!,
    now: () => now,
    minimizedTrackingLossGraceMs: 500,
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(tracker.isTracking(), true);

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(tracker.isTracking(), true);
  assert.equal(tracker.isTargetWindowMinimized(), true);
  assert.deepEqual(tracker.getGeometry(), { x: 10, y: 20, width: 1280, height: 720 });

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(tracker.isTracking(), true);
  assert.equal(tracker.isTargetWindowMinimized(), true);
  assert.deepEqual(tracker.getGeometry(), { x: 10, y: 20, width: 1280, height: 720 });

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(tracker.isTracking(), true);
  assert.equal(tracker.isTargetWindowMinimized(), false);
  assert.deepEqual(tracker.getGeometry(), { x: 10, y: 20, width: 1280, height: 720 });
});
