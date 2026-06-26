import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { getDefaultLauncherLogFile, getDefaultMpvLogFile } from './types.js';

function localDateKey(date: Date): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(
    date.getDate(),
  ).padStart(2, '0')}`;
}

test('getDefaultMpvLogFile uses APPDATA on windows', () => {
  const today = localDateKey(new Date());
  const resolved = getDefaultMpvLogFile({
    platform: 'win32',
    homeDir: 'C:\\Users\\tester',
    appDataDir: 'C:\\Users\\tester\\AppData\\Roaming',
  });

  assert.equal(
    path.normalize(resolved),
    path.normalize(
      path.join('C:\\Users\\tester\\AppData\\Roaming', 'SubMiner', 'logs', `mpv-${today}.log`),
    ),
  );
});

test('getDefaultLauncherLogFile uses launcher prefix', () => {
  const today = localDateKey(new Date());
  const resolved = getDefaultLauncherLogFile({
    platform: 'linux',
    homeDir: '/home/tester',
  });

  assert.equal(
    resolved,
    path.join('/home/tester', '.config', 'SubMiner', 'logs', `launcher-${today}.log`),
  );
});
