import test from 'node:test';
import assert from 'node:assert/strict';
import {
  applyWindowsMpvShortcuts,
  buildWindowsMpvShortcutDetails,
  detectWindowsMpvShortcuts,
  resolveWindowsMpvShortcutPaths,
  resolveWindowsStartMenuProgramsDir,
} from './windows-mpv-shortcuts';

test('resolveWindowsStartMenuProgramsDir derives Programs folder from APPDATA', () => {
  assert.equal(
    resolveWindowsStartMenuProgramsDir('C:\\Users\\tester\\AppData\\Roaming'),
    'C:\\Users\\tester\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs',
  );
});

test('resolveWindowsMpvShortcutPaths builds start menu and desktop lnk paths', () => {
  const paths = resolveWindowsMpvShortcutPaths({
    appDataDir: 'C:\\Users\\tester\\AppData\\Roaming',
    desktopDir: 'C:\\Users\\tester\\Desktop',
  });

  assert.equal(
    paths.startMenuPath,
    'C:\\Users\\tester\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\SubMiner mpv.lnk',
  );
  assert.equal(paths.desktopPath, 'C:\\Users\\tester\\Desktop\\SubMiner mpv.lnk');
});

test('buildWindowsMpvShortcutDetails targets SubMiner.exe with --launch-mpv', () => {
  assert.deepEqual(buildWindowsMpvShortcutDetails('C:\\Apps\\SubMiner\\SubMiner.exe'), {
    target: 'C:\\Apps\\SubMiner\\SubMiner.exe',
    args: '--launch-mpv',
    cwd: 'C:\\Apps\\SubMiner',
    description: 'Launch mpv with the SubMiner profile',
    icon: 'C:\\Apps\\SubMiner\\SubMiner.exe',
    iconIndex: 0,
  });
});

test('detectWindowsMpvShortcuts reflects existing shortcuts', () => {
  const detected = detectWindowsMpvShortcuts(
    {
      startMenuPath: 'C:\\Programs\\SubMiner mpv.lnk',
      desktopPath: 'C:\\Desktop\\SubMiner mpv.lnk',
    },
    (candidate) => candidate === 'C:\\Desktop\\SubMiner mpv.lnk',
  );

  assert.deepEqual(detected, {
    startMenuInstalled: false,
    desktopInstalled: true,
  });
});

test('applyWindowsMpvShortcuts creates enabled shortcuts and removes disabled ones', () => {
  const writes: string[] = [];
  const removes: string[] = [];
  const result = applyWindowsMpvShortcuts({
    preferences: {
      startMenuEnabled: true,
      desktopEnabled: false,
    },
    paths: {
      startMenuPath: 'C:\\Programs\\SubMiner mpv.lnk',
      desktopPath: 'C:\\Desktop\\SubMiner mpv.lnk',
    },
    exePath: 'C:\\Apps\\SubMiner\\SubMiner.exe',
    writeShortcutLink: (shortcutPath, _operation, details) => {
      writes.push(`${shortcutPath}|${details.target}|${details.args}`);
      return true;
    },
    rmSync: (candidate) => {
      removes.push(candidate);
    },
    mkdirSync: () => undefined,
  });

  assert.equal(result.ok, true);
  assert.equal(result.status, 'installed');
  assert.deepEqual(writes, [
    'C:\\Programs\\SubMiner mpv.lnk|C:\\Apps\\SubMiner\\SubMiner.exe|--launch-mpv',
  ]);
  assert.deepEqual(removes, ['C:\\Desktop\\SubMiner mpv.lnk']);
});

test('applyWindowsMpvShortcuts returns skipped when both shortcuts are disabled', () => {
  const removes: string[] = [];
  const result = applyWindowsMpvShortcuts({
    preferences: {
      startMenuEnabled: false,
      desktopEnabled: false,
    },
    paths: {
      startMenuPath: 'C:\\Programs\\SubMiner mpv.lnk',
      desktopPath: 'C:\\Desktop\\SubMiner mpv.lnk',
    },
    exePath: 'C:\\Apps\\SubMiner\\SubMiner.exe',
    writeShortcutLink: () => true,
    rmSync: (candidate) => {
      removes.push(candidate);
    },
    mkdirSync: () => undefined,
  });

  assert.equal(result.ok, true);
  assert.equal(result.status, 'skipped');
  assert.deepEqual(removes, ['C:\\Programs\\SubMiner mpv.lnk', 'C:\\Desktop\\SubMiner mpv.lnk']);
});

test('applyWindowsMpvShortcuts reports write failures', () => {
  const result = applyWindowsMpvShortcuts({
    preferences: {
      startMenuEnabled: true,
      desktopEnabled: true,
    },
    paths: {
      startMenuPath: 'C:\\Programs\\SubMiner mpv.lnk',
      desktopPath: 'C:\\Desktop\\SubMiner mpv.lnk',
    },
    exePath: 'C:\\Apps\\SubMiner\\SubMiner.exe',
    writeShortcutLink: (shortcutPath) => shortcutPath.endsWith('Desktop\\SubMiner mpv.lnk'),
    mkdirSync: () => undefined,
  });

  assert.equal(result.ok, false);
  assert.equal(result.status, 'failed');
  assert.match(result.message, /C:\\Programs\\SubMiner mpv\.lnk/);
});
