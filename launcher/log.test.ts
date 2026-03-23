import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { getDefaultLauncherLogFile, getDefaultMpvLogFile } from './types.js';

test('getDefaultMpvLogFile uses APPDATA on windows', () => {
  const resolved = getDefaultMpvLogFile({
    platform: 'win32',
    homeDir: 'C:\\Users\\tester',
    appDataDir: 'C:\\Users\\tester\\AppData\\Roaming',
  });

  assert.equal(
    path.normalize(resolved),
    path.normalize(
      path.join(
        'C:\\Users\\tester\\AppData\\Roaming',
        'SubMiner',
        'logs',
        `mpv-${new Date().toISOString().slice(0, 10)}.log`,
      ),
    ),
  );
});

test('getDefaultLauncherLogFile uses launcher prefix', () => {
  const resolved = getDefaultLauncherLogFile({
    platform: 'linux',
    homeDir: '/home/tester',
  });

  assert.equal(
    resolved,
    path.join(
      '/home/tester',
      '.config',
      'SubMiner',
      'logs',
      `launcher-${new Date().toISOString().slice(0, 10)}.log`,
    ),
  );
});
