import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, utimesSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import {
  isCompiledMacOSHelperCurrent,
  MacOSWindowTracker,
  parseMacOSHelperOutput,
} from './macos-tracker';

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

test('parseMacOSHelperOutput parses inactive state without geometry', () => {
  assert.deepEqual(parseMacOSHelperOutput('inactive'), {
    geometry: null,
    focused: false,
    inactive: true,
  });
});

test('isCompiledMacOSHelperCurrent rejects binaries older than the Swift source', () => {
  const tempDir = mkdtempSync(join(tmpdir(), 'subminer-macos-helper-'));
  try {
    const binaryPath = join(tempDir, 'get-mpv-window-macos');
    const sourcePath = join(tempDir, 'get-mpv-window-macos.swift');
    writeFileSync(binaryPath, 'binary');
    writeFileSync(sourcePath, 'source');

    const older = new Date('2026-01-01T00:00:00Z');
    const newer = new Date('2026-01-01T00:00:05Z');
    utimesSync(binaryPath, older, older);
    utimesSync(sourcePath, newer, newer);

    assert.equal(isCompiledMacOSHelperCurrent(binaryPath, sourcePath), false);

    utimesSync(binaryPath, newer, newer);
    utimesSync(sourcePath, older, older);

    assert.equal(isCompiledMacOSHelperCurrent(binaryPath, sourcePath), true);
  } finally {
    rmSync(tempDir, { recursive: true, force: true });
  }
});

test('MacOSWindowTracker slows polling while focused target is stable', async () => {
  const scheduledDelays: number[] = [];
  let callIndex = 0;
  const tracker = new MacOSWindowTracker('/tmp/mpv.sock', {
    resolveHelper: () => ({
      helperPath: 'helper',
      helperType: 'binary',
    }),
    runHelper: async () => {
      callIndex += 1;
      return { stdout: '10,20,1280,720,1', stderr: '' };
    },
    fastPollIntervalMs: 250,
    stablePollIntervalMs: 1_000,
    setPollTimeout: ((_callback: () => void, delayMs: number) => {
      scheduledDelays.push(delayMs);
      return {} as ReturnType<typeof setTimeout>;
    }) as never,
    clearPollTimeout: (() => {}) as never,
  } as never);

  tracker.start();
  await new Promise((resolve) => setTimeout(resolve, 0));
  tracker.stop();

  assert.equal(callIndex, 1);
  assert.deepEqual(scheduledDelays, [1_000]);
});

test('MacOSWindowTracker keeps fast polling while target is not focused', async () => {
  const scheduledDelays: number[] = [];
  const tracker = new MacOSWindowTracker('/tmp/mpv.sock', {
    resolveHelper: () => ({
      helperPath: 'helper',
      helperType: 'binary',
    }),
    runHelper: async () => ({ stdout: '10,20,1280,720,0', stderr: '' }),
    fastPollIntervalMs: 250,
    stablePollIntervalMs: 1_000,
    setPollTimeout: ((_callback: () => void, delayMs: number) => {
      scheduledDelays.push(delayMs);
      return {} as ReturnType<typeof setTimeout>;
    }) as never,
    clearPollTimeout: (() => {}) as never,
  } as never);

  tracker.start();
  await new Promise((resolve) => setTimeout(resolve, 0));
  tracker.stop();

  assert.deepEqual(scheduledDelays, [250]);
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

test('MacOSWindowTracker preserves target focus on helper not-found while retaining geometry', async () => {
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

test('MacOSWindowTracker keeps previously focused target through repeated not-found misses after grace', async () => {
  let callIndex = 0;
  let now = 1_000;
  const focusChanges: boolean[] = [];
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
  tracker.onWindowFocusChange = (focused) => {
    focusChanges.push(focused);
  };

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.isTracking(), true);
  assert.equal(tracker.isTargetWindowFocused(), true);

  now += 1_000;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));

  now += 1_000;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));

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
  assert.deepEqual(focusChanges, [true]);
});

test('MacOSWindowTracker keeps previously focused target through repeated helper execution failures', async () => {
  let callIndex = 0;
  let now = 1_000;
  const focusChanges: boolean[] = [];

  const tracker = new MacOSWindowTracker('/tmp/mpv.sock', {
    resolveHelper: () => ({
      helperPath: 'helper.swift',
      helperType: 'swift',
    }),
    runHelper: async () => {
      callIndex += 1;
      if (callIndex === 1) {
        return { stdout: '10,20,1280,720,1', stderr: '' };
      }
      throw Object.assign(new Error('helper timed out'), { stderr: 'timeout' });
    },
    now: () => now,
    trackingLossGraceMs: 500,
  });
  tracker.onWindowFocusChange = (focused) => {
    focusChanges.push(focused);
  };

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));

  now += 1_000;
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));

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
  assert.deepEqual(focusChanges, [true]);
});

test('MacOSWindowTracker marks target unfocused on explicit inactive helper signal', async () => {
  let callIndex = 0;
  const focusChanges: boolean[] = [];
  const outputs = [
    { stdout: '10,20,1280,720,1', stderr: '' },
    { stdout: 'inactive', stderr: '' },
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

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.equal(tracker.isTracking(), true);
  assert.deepEqual(tracker.getGeometry(), {
    x: 10,
    y: 20,
    width: 1280,
    height: 720,
  });
  assert.equal(tracker.isTargetWindowFocused(), false);
  assert.deepEqual(focusChanges, [true, false]);
});

test('MacOSWindowTracker drops tracking after consecutive helper misses', async () => {
  let callIndex = 0;
  const outputs = [
    { stdout: '10,20,1280,720,0', stderr: '' },
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
  assert.equal(tracker.isTargetWindowFocused(), false);

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
    { stdout: '10,20,1280,720,0', stderr: '' },
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
  assert.equal(tracker.isTargetWindowFocused(), false);

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
