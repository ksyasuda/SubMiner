import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { getDefaultLauncherLogFile, getDefaultMpvLogFile } from './types.js';

test('getDefaultMpvLogFile uses APPDATA on windows', () => {
  const today = new Date().toISOString().slice(0, 10);
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
        `mpv-${today}.log`,
      ),
    ),
  );
});

test('getDefaultLauncherLogFile uses launcher prefix', () => {
  const today = new Date().toISOString().slice(0, 10);
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
      `launcher-${today}.log`,
    ),
  );
});
