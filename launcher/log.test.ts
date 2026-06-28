import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { getDefaultLauncherLogFile, getDefaultMpvLogFile } from './types.js';
import { localDateKey } from '../src/shared/log-files.js';

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
