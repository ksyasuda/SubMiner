import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { resolveDefaultLogFilePath, setLogRotation } from './logger';

test('resolveDefaultLogFilePath uses APPDATA on windows', () => {
  const today = new Date().toISOString().slice(0, 10);
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
        `app-${today}.log`,
      ),
    ),
  );
});

test('resolveDefaultLogFilePath uses .config on linux', () => {
  const today = new Date().toISOString().slice(0, 10);
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
      `app-${today}.log`,
    ),
  );
});

test('setLogRotation accepts numeric retention days', () => {
  setLogRotation(14);
  try {
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
    assert.equal(process.env.SUBMINER_LOG_ROTATION, '14');
  } finally {
    setLogRotation(7);
  }
});
