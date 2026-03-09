import test from 'node:test';
import assert from 'node:assert/strict';
import { WindowsWindowTracker } from './windows-tracker';

test('WindowsWindowTracker skips overlapping polls while helper is in flight', async () => {
  let helperCalls = 0;
  let release: (() => void) | undefined;
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });

  const tracker = new WindowsWindowTracker(undefined, {
    resolveHelper: () => ({
      kind: 'powershell',
      command: 'powershell.exe',
      args: ['-File', 'helper.ps1'],
      helperPath: 'helper.ps1',
    }),
    runHelper: async () => {
      helperCalls += 1;
      await gate;
      return {
        stdout: '0,0,640,360',
        stderr: 'focus=focused',
      };
    },
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(helperCalls, 1);

  assert.ok(release);
  release();
  await new Promise((resolve) => setTimeout(resolve, 0));
});

test('WindowsWindowTracker updates geometry from helper output', async () => {
  const tracker = new WindowsWindowTracker(undefined, {
    resolveHelper: () => ({
      kind: 'powershell',
      command: 'powershell.exe',
      args: ['-File', 'helper.ps1'],
      helperPath: 'helper.ps1',
    }),
    runHelper: async () => ({
      stdout: '10,20,1280,720',
      stderr: 'focus=focused',
    }),
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(tracker.getGeometry(), {
    x: 10,
    y: 20,
    width: 1280,
    height: 720,
  });
  assert.equal(tracker.isTargetWindowFocused(), true);
});

test('WindowsWindowTracker clears geometry for helper misses', async () => {
  const tracker = new WindowsWindowTracker(undefined, {
    resolveHelper: () => ({
      kind: 'powershell',
      command: 'powershell.exe',
      args: ['-File', 'helper.ps1'],
      helperPath: 'helper.ps1',
    }),
    runHelper: async () => ({
      stdout: 'not-found',
      stderr: 'focus=not-focused',
    }),
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tracker.getGeometry(), null);
  assert.equal(tracker.isTargetWindowFocused(), false);
});

test('WindowsWindowTracker retries without socket filter when filtered helper lookup misses', async () => {
  const helperCalls: Array<string | null> = [];
  const tracker = new WindowsWindowTracker('\\\\.\\pipe\\subminer-socket', {
    resolveHelper: () => ({
      kind: 'powershell',
      command: 'powershell.exe',
      args: ['-File', 'helper.ps1'],
      helperPath: 'helper.ps1',
    }),
    runHelper: async (_spec, _mode, targetMpvSocketPath) => {
      helperCalls.push(targetMpvSocketPath);
      if (targetMpvSocketPath) {
        return {
          stdout: 'not-found',
          stderr: 'focus=not-focused',
        };
      }
      return {
        stdout: '25,30,1440,810',
        stderr: 'focus=focused',
      };
    },
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(helperCalls, ['\\\\.\\pipe\\subminer-socket', null]);
  assert.deepEqual(tracker.getGeometry(), {
    x: 25,
    y: 30,
    width: 1440,
    height: 810,
  });
  assert.equal(tracker.isTargetWindowFocused(), true);
});
