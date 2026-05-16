import assert from 'node:assert/strict';
import test from 'node:test';
import { MacOSWindowTracker, parseMacOSHelperOutput } from './macos-tracker';

test('parseMacOSHelperOutput parses minimized state', () => {
  assert.deepEqual(parseMacOSHelperOutput('minimized'), {
    geometry: null,
    focused: false,
    minimized: true,
  });
});

test('parseMacOSHelperOutput parses active focused state without geometry', () => {
  assert.deepEqual(parseMacOSHelperOutput('active'), {
    geometry: null,
    focused: true,
    active: true,
  });
});

test('MacOSWindowTracker keeps the last geometry through a single helper miss', async () => {
  let callIndex = 0;
  const outputs = [
    { stdout: '10,20,1280,720,1', stderr: '' },
    { stdout: 'not-found', stderr: '' },
    { stdout: '10,20,1280,720,1', stderr: '' },
  ];

  const tracker = new MacOSWindowTracker('/tmp/mpv.sock', {
    resolveHelper: () => ({
      helperPath: 'helper.swift',
      helperType: 'swift',
    }),
    runHelper: async () => outputs[callIndex++] ?? outputs.at(-1)!,
    trackingLossGraceMs: 0,
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.deepEqual(tracker.getGeometry(), {
    x: 10,
    y: 20,
    width: 1280,
    height: 720,
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.deepEqual(tracker.getGeometry(), {
    x: 10,
    y: 20,
    width: 1280,
    height: 720,
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.deepEqual(tracker.getGeometry(), {
    x: 10,
    y: 20,
    width: 1280,
    height: 720,
  });
});

test('MacOSWindowTracker preserves target focus during transient helper misses', async () => {
  let callIndex = 0;
  const focusChanges: boolean[] = [];
  const outputs = [
    { stdout: '10,20,1280,720,1', stderr: '' },
    { stdout: 'not-found', stderr: '' },
  ];

  const tracker = new MacOSWindowTracker('/tmp/mpv.sock', {
    resolveHelper: () => ({
      helperPath: 'helper.swift',
      helperType: 'swift',
    }),
    runHelper: async () => outputs[callIndex++] ?? outputs.at(-1)!,
    trackingLossGraceMs: 1_500,
  });
  tracker.onWindowFocusChange = (focused) => {
    focusChanges.push(focused);
  };

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTargetWindowFocused(), true);

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.equal(tracker.isTracking(), true);
  assert.deepEqual(tracker.getGeometry(), {
    x: 10,
    y: 20,
    width: 1280,
    height: 720,
  });
  assert.equal(tracker.isTargetWindowFocused(), true);
  assert.deepEqual(focusChanges, [true]);
});

test('MacOSWindowTracker keeps focused fullscreen target through active helper misses after grace', async () => {
  let callIndex = 0;
  let now = 1_000;
  const outputs = [
    { stdout: '10,20,1280,720,1', stderr: '' },
    { stdout: 'active', stderr: '' },
    { stdout: 'active', stderr: '' },
  ];

  const tracker = new MacOSWindowTracker('/tmp/mpv.sock', {
    resolveHelper: () => ({
      helperPath: 'helper.swift',
      helperType: 'swift',
    }),
    runHelper: async () => outputs[callIndex++] ?? outputs.at(-1)!,
    now: () => now,
    trackingLossGraceMs: 500,
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTracking(), true);
  assert.equal(tracker.isTargetWindowFocused(), true);

  now += 1_000;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTracking(), true);

  now += 1_000;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.equal(tracker.isTracking(), true);
  assert.equal(tracker.isTargetWindowFocused(), true);
  assert.deepEqual(tracker.getGeometry(), {
    x: 10,
    y: 20,
    width: 1280,
    height: 720,
  });
});

test('MacOSWindowTracker drops tracking after consecutive helper misses', async () => {
  let callIndex = 0;
  const outputs = [
    { stdout: '10,20,1280,720,1', stderr: '' },
    { stdout: 'not-found', stderr: '' },
    { stdout: 'not-found', stderr: '' },
  ];

  const tracker = new MacOSWindowTracker('/tmp/mpv.sock', {
    resolveHelper: () => ({
      helperPath: 'helper.swift',
      helperType: 'swift',
    }),
    runHelper: async () => outputs[callIndex++] ?? outputs.at(-1)!,
    trackingLossGraceMs: 0,
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTracking(), true);

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTracking(), true);

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTracking(), false);
  assert.equal(tracker.getGeometry(), null);
  assert.equal(tracker.isTargetWindowFocused(), false);
});

test('MacOSWindowTracker keeps tracking through repeated helper misses inside grace window', async () => {
  let callIndex = 0;
  let now = 1_000;
  const outputs = [
    { stdout: '10,20,1280,720,1', stderr: '' },
    { stdout: 'not-found', stderr: '' },
    { stdout: 'not-found', stderr: '' },
    { stdout: 'not-found', stderr: '' },
    { stdout: 'not-found', stderr: '' },
  ];

  const tracker = new MacOSWindowTracker('/tmp/mpv.sock', {
    resolveHelper: () => ({
      helperPath: 'helper.swift',
      helperType: 'swift',
    }),
    runHelper: async () => outputs[callIndex++] ?? outputs.at(-1)!,
    now: () => now,
    trackingLossGraceMs: 1_500,
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTracking(), true);

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTracking(), true);

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTracking(), true);

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTracking(), true);
  assert.deepEqual(tracker.getGeometry(), {
    x: 10,
    y: 20,
    width: 1280,
    height: 720,
  });
});

test('MacOSWindowTracker drops tracking after grace window expires', async () => {
  let callIndex = 0;
  let now = 1_000;
  const outputs = [
    { stdout: '10,20,1280,720,1', stderr: '' },
    { stdout: 'not-found', stderr: '' },
    { stdout: 'not-found', stderr: '' },
    { stdout: 'not-found', stderr: '' },
  ];

  const tracker = new MacOSWindowTracker('/tmp/mpv.sock', {
    resolveHelper: () => ({
      helperPath: 'helper.swift',
      helperType: 'swift',
    }),
    runHelper: async () => outputs[callIndex++] ?? outputs.at(-1)!,
    now: () => now,
    trackingLossGraceMs: 500,
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTracking(), true);

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTracking(), true);

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTracking(), true);

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTracking(), true);

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTracking(), false);
  assert.equal(tracker.getGeometry(), null);
});

test('MacOSWindowTracker reports minimized target when helper reports minimized', async () => {
  let callIndex = 0;
  let now = 1_000;
  const outputs = [
    { stdout: '10,20,1280,720,1', stderr: '' },
    { stdout: 'minimized', stderr: '' },
    { stdout: 'minimized', stderr: '' },
  ];

  const tracker = new MacOSWindowTracker('/tmp/mpv.sock', {
    resolveHelper: () => ({
      helperPath: 'helper.swift',
      helperType: 'swift',
    }),
    runHelper: async () => outputs[callIndex++] ?? outputs.at(-1)!,
    now: () => now,
    minimizedTrackingLossGraceMs: 200,
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTracking(), true);
  assert.equal(tracker.isTargetWindowMinimized(), false);

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTargetWindowMinimized(), true);
  assert.equal(tracker.isTracking(), true);

  now += 250;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTargetWindowMinimized(), true);
  assert.equal(tracker.isTracking(), false);
});
