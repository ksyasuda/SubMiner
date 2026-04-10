import test from 'node:test';
import assert from 'node:assert/strict';
import {
  lowerWindowsOverlayInZOrder,
  parseWindowTrackerHelperForegroundProcess,
  parseWindowTrackerHelperFocusState,
  parseWindowTrackerHelperOutput,
  parseWindowTrackerHelperState,
  queryWindowsForegroundProcessName,
  queryWindowsTargetWindowHandle,
  queryWindowsTrackerMpvWindows,
  resolveWindowsTrackerHelper,
  syncWindowsOverlayToMpvZOrder,
} from './windows-helper';

test('parseWindowTrackerHelperOutput parses helper geometry output', () => {
  assert.deepEqual(parseWindowTrackerHelperOutput('120,240,1280,720'), {
    x: 120,
    y: 240,
    width: 1280,
    height: 720,
  });
});

test('parseWindowTrackerHelperOutput returns null for misses and invalid payloads', () => {
  assert.equal(parseWindowTrackerHelperOutput('not-found'), null);
  assert.equal(parseWindowTrackerHelperOutput('1,2,3'), null);
  assert.equal(parseWindowTrackerHelperOutput('1,2,0,4'), null);
});

test('parseWindowTrackerHelperFocusState parses helper stderr metadata', () => {
  assert.equal(parseWindowTrackerHelperFocusState('focus=focused'), true);
  assert.equal(parseWindowTrackerHelperFocusState('focus=not-focused'), false);
  assert.equal(parseWindowTrackerHelperFocusState('warning\nfocus=focused\nnote'), true);
  assert.equal(parseWindowTrackerHelperFocusState(''), null);
});

test('parseWindowTrackerHelperState parses helper stderr metadata', () => {
  assert.equal(parseWindowTrackerHelperState('state=visible'), 'visible');
  assert.equal(parseWindowTrackerHelperState('focus=not-focused\nstate=minimized'), 'minimized');
  assert.equal(parseWindowTrackerHelperState('state=unknown'), null);
  assert.equal(parseWindowTrackerHelperState(''), null);
});

test('parseWindowTrackerHelperForegroundProcess parses helper stdout metadata', () => {
  assert.equal(parseWindowTrackerHelperForegroundProcess('process=mpv'), 'mpv');
  assert.equal(parseWindowTrackerHelperForegroundProcess('process=chrome'), 'chrome');
  assert.equal(parseWindowTrackerHelperForegroundProcess('not-found'), null);
  assert.equal(parseWindowTrackerHelperForegroundProcess(''), null);
});

test('queryWindowsForegroundProcessName reads foreground process from powershell helper', async () => {
  const processName = await queryWindowsForegroundProcessName({
    resolveHelper: () => ({
      kind: 'powershell',
      command: 'powershell.exe',
      args: ['-File', 'helper.ps1'],
      helperPath: 'helper.ps1',
    }),
    runHelper: async () => ({
      stdout: 'process=mpv',
      stderr: '',
    }),
  });

  assert.equal(processName, 'mpv');
});

test('queryWindowsForegroundProcessName returns null when no powershell helper is available', async () => {
  const processName = await queryWindowsForegroundProcessName({
    resolveHelper: () => ({
      kind: 'native',
      command: 'helper.exe',
      args: [],
      helperPath: 'helper.exe',
    }),
  });

  assert.equal(processName, null);
});

test('syncWindowsOverlayToMpvZOrder forwards socket path and overlay handle to powershell helper', async () => {
  let capturedMode: string | null = null;
  let capturedArgs: string[] | null = null;

  const synced = await syncWindowsOverlayToMpvZOrder({
    overlayWindowHandle: '12345',
    targetMpvSocketPath: '\\\\.\\pipe\\subminer-socket',
    resolveHelper: () => ({
      kind: 'powershell',
      command: 'powershell.exe',
      args: ['-File', 'helper.ps1'],
      helperPath: 'helper.ps1',
    }),
    runHelper: async (_spec, mode, extraArgs = []) => {
      capturedMode = mode;
      capturedArgs = extraArgs;
      return {
        stdout: 'ok',
        stderr: '',
      };
    },
  });

  assert.equal(synced, true);
  assert.equal(capturedMode, 'bind-overlay');
  assert.deepEqual(capturedArgs, ['\\\\.\\pipe\\subminer-socket', '12345']);
});

test('lowerWindowsOverlayInZOrder forwards overlay handle to powershell helper', async () => {
  let capturedMode: string | null = null;
  let capturedArgs: string[] | null = null;

  const lowered = await lowerWindowsOverlayInZOrder({
    overlayWindowHandle: '67890',
    resolveHelper: () => ({
      kind: 'powershell',
      command: 'powershell.exe',
      args: ['-File', 'helper.ps1'],
      helperPath: 'helper.ps1',
    }),
    runHelper: async (_spec, mode, extraArgs = []) => {
      capturedMode = mode;
      capturedArgs = extraArgs;
      return {
        stdout: 'ok',
        stderr: '',
      };
    },
  });

  assert.equal(lowered, true);
  assert.equal(capturedMode, 'lower-overlay');
  assert.deepEqual(capturedArgs, ['', '67890']);
});

test('queryWindowsTrackerMpvWindows resolves geometry from the powershell helper', () => {
  let capturedMode: string | null = null;
  let capturedArgs: string[] | null = null;

  const result = queryWindowsTrackerMpvWindows({
    targetMpvSocketPath: '\\\\.\\pipe\\subminer-socket',
    resolveHelper: () => ({
      kind: 'powershell',
      command: 'powershell.exe',
      args: ['-File', 'helper.ps1'],
      helperPath: 'helper.ps1',
    }),
    runHelperSync: (_spec, mode, extraArgs = []) => {
      capturedMode = mode;
      capturedArgs = extraArgs;
      return {
        stdout: '120,240,1280,720',
        stderr: 'focus=focused\nstate=visible',
      };
    },
  });

  assert.deepEqual(result, {
    matches: [
      {
        hwnd: 0,
        bounds: {
          x: 120,
          y: 240,
          width: 1280,
          height: 720,
        },
        area: 1280 * 720,
        isForeground: true,
      },
    ],
    focusState: true,
    windowState: 'visible',
  });
  assert.equal(capturedMode, 'geometry');
  assert.deepEqual(capturedArgs, ['\\\\.\\pipe\\subminer-socket']);
});

test('queryWindowsTargetWindowHandle resolves the selected hwnd from the powershell helper', () => {
  let capturedMode: string | null = null;
  let capturedArgs: string[] | null = null;

  const hwnd = queryWindowsTargetWindowHandle({
    targetMpvSocketPath: '\\\\.\\pipe\\subminer-socket',
    resolveHelper: () => ({
      kind: 'powershell',
      command: 'powershell.exe',
      args: ['-File', 'helper.ps1'],
      helperPath: 'helper.ps1',
    }),
    runHelperSync: (_spec, mode, extraArgs = []) => {
      capturedMode = mode;
      capturedArgs = extraArgs;
      return {
        stdout: '12345',
        stderr: '',
      };
    },
  });

  assert.equal(hwnd, 12345);
  assert.equal(capturedMode, 'target-hwnd');
  assert.deepEqual(capturedArgs, ['\\\\.\\pipe\\subminer-socket']);
});

test('resolveWindowsTrackerHelper auto mode prefers native helper when present', () => {
  const helper = resolveWindowsTrackerHelper({
    dirname: 'C:\\repo\\dist\\window-trackers',
    resourcesPath: 'C:\\repo\\resources',
    existsSync: (candidate) =>
      candidate === 'C:\\repo\\resources\\scripts\\get-mpv-window-windows.exe',
    helperModeEnv: 'auto',
  });

  assert.deepEqual(helper, {
    kind: 'native',
    command: 'C:\\repo\\resources\\scripts\\get-mpv-window-windows.exe',
    args: [],
    helperPath: 'C:\\repo\\resources\\scripts\\get-mpv-window-windows.exe',
  });
});

test('resolveWindowsTrackerHelper auto mode falls back to powershell helper', () => {
  const helper = resolveWindowsTrackerHelper({
    dirname: 'C:\\repo\\dist\\window-trackers',
    resourcesPath: 'C:\\repo\\resources',
    existsSync: (candidate) =>
      candidate === 'C:\\repo\\resources\\scripts\\get-mpv-window-windows.ps1',
    helperModeEnv: 'auto',
  });

  assert.deepEqual(helper, {
    kind: 'powershell',
    command: 'powershell.exe',
    args: [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      'C:\\repo\\resources\\scripts\\get-mpv-window-windows.ps1',
    ],
    helperPath: 'C:\\repo\\resources\\scripts\\get-mpv-window-windows.ps1',
  });
});

test('resolveWindowsTrackerHelper explicit powershell mode ignores native helper', () => {
  const helper = resolveWindowsTrackerHelper({
    dirname: 'C:\\repo\\dist\\window-trackers',
    resourcesPath: 'C:\\repo\\resources',
    existsSync: (candidate) =>
      candidate === 'C:\\repo\\resources\\scripts\\get-mpv-window-windows.exe' ||
      candidate === 'C:\\repo\\resources\\scripts\\get-mpv-window-windows.ps1',
    helperModeEnv: 'powershell',
  });

  assert.equal(helper?.kind, 'powershell');
  assert.equal(helper?.helperPath, 'C:\\repo\\resources\\scripts\\get-mpv-window-windows.ps1');
});

test('resolveWindowsTrackerHelper explicit native mode fails cleanly when helper is missing', () => {
  const helper = resolveWindowsTrackerHelper({
    dirname: 'C:\\repo\\dist\\window-trackers',
    resourcesPath: 'C:\\repo\\resources',
    existsSync: () => false,
    helperModeEnv: 'native',
  });

  assert.equal(helper, null);
});

test('resolveWindowsTrackerHelper explicit helper path overrides default search', () => {
  const helper = resolveWindowsTrackerHelper({
    dirname: 'C:\\repo\\dist\\window-trackers',
    resourcesPath: 'C:\\repo\\resources',
    existsSync: (candidate) => candidate === 'D:\\custom\\tracker.ps1',
    helperModeEnv: 'auto',
    helperPathEnv: 'D:\\custom\\tracker.ps1',
  });

  assert.deepEqual(helper, {
    kind: 'powershell',
    command: 'powershell.exe',
    args: ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', 'D:\\custom\\tracker.ps1'],
    helperPath: 'D:\\custom\\tracker.ps1',
  });
});
