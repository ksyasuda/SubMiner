import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createOverlayVisibilityDiagnostics,
  type OverlayVisibilitySnapshot,
} from './overlay-visibility-diagnostics';

function harness() {
  let now = 0;
  let callback: (() => void) | null = null;
  let delay = 0;
  const warnings: Array<{ message: string; details: Record<string, unknown> }> = [];
  let snapshot: OverlayVisibilitySnapshot | null = {
    expectedVisible: true,
    visible: false,
    opacity: 0,
    contentReady: false,
    tracking: false,
    targetPid: 4242,
    pidLookupPending: false,
  };
  const diagnostics = createOverlayVisibilityDiagnostics({
    now: () => now,
    warn: (message, details) => warnings.push({ message, details }),
    schedule: (fn, ms) => {
      callback = fn;
      delay = ms;
      return { unref() {} } as ReturnType<typeof setTimeout>;
    },
    cancel: () => {
      callback = null;
    },
  });
  const update = () => diagnostics.update(() => snapshot);
  return {
    warnings,
    update,
    setSnapshot: (value: typeof snapshot) => {
      snapshot = value;
    },
    tick: () => {
      assert.ok(callback, 'watchdog must be scheduled');
      now += delay;
      callback();
    },
    scheduled: () => callback !== null,
  };
}

test('visibility watchdog warns without more events and rate limits repeated polls', () => {
  const h = harness();
  h.update();
  assert.equal(h.warnings.length, 0);
  for (let i = 0; i < 100; i += 1) h.update();
  h.tick();
  assert.equal(h.warnings.length, 1);
  assert.equal(h.warnings[0]?.details.reason, 'mpv-window-not-tracked');
  assert.equal(h.warnings[0]?.details.targetPid, 4242);
  assert.equal(h.warnings[0]?.details.elapsedMs, 5_000);
  for (let i = 0; i < 100; i += 1) h.update();
  assert.equal(h.warnings.length, 1);
  h.tick();
  assert.equal(h.warnings[1]?.details.elapsedMs, 35_000);
});

test('visibility watchdog distinguishes renderer readiness, hidden windows and opacity', () => {
  for (const [contentReady, visible, opacity, reason] of [
    [false, false, 0, 'renderer-not-ready'],
    [true, false, 0, 'window-hidden'],
    [true, true, 0, 'window-transparent'],
  ] as const) {
    const h = harness();
    h.setSnapshot({ expectedVisible: true, tracking: true, contentReady, visible, opacity });
    h.update();
    h.tick();
    assert.equal(h.warnings[0]?.details.reason, reason);
  }
});

test('visibility watchdog reports recovery only after warning and stops', () => {
  const h = harness();
  h.update();
  h.tick();
  h.setSnapshot({
    expectedVisible: true,
    visible: true,
    opacity: 1,
    contentReady: true,
    tracking: true,
  });
  h.tick();
  assert.match(h.warnings[1]!.message, /recovered/);
  assert.equal(h.scheduled(), false);
  h.update();
  assert.equal(h.warnings.length, 2);
});

test('visibility watchdog stays quiet after fast startup, intentional hide or window destruction', () => {
  for (const snapshot of [
    { expectedVisible: true, visible: true, opacity: 1, contentReady: true, tracking: true },
    { expectedVisible: false, visible: false, opacity: 0, contentReady: false, tracking: false },
    null,
  ]) {
    const h = harness();
    h.update();
    h.setSnapshot(snapshot);
    h.tick();
    assert.deepEqual(h.warnings, []);
    assert.equal(h.scheduled(), false);
  }
});
