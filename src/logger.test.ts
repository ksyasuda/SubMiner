import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { resolveDefaultLogFilePath } from './logger';

test('resolveDefaultLogFilePath uses APPDATA on windows', () => {
  const resolved = resolveDefaultLogFilePath({
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
        `app-${new Date().toISOString().slice(0, 10)}.log`,
      ),
    ),
  );
});

test('resolveDefaultLogFilePath uses .config on linux', () => {
  const resolved = resolveDefaultLogFilePath({
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
      `app-${new Date().toISOString().slice(0, 10)}.log`,
    ),
  );
});
