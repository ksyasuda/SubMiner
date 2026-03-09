import test from 'node:test';
import assert from 'node:assert/strict';
import {
  parseWindowTrackerHelperFocusState,
  parseWindowTrackerHelperOutput,
  resolveWindowsTrackerHelper,
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
