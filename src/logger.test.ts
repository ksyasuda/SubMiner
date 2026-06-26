import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { resolveDefaultLogFilePath, setLogRotation } from './logger';

function localDateKey(date: Date): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(
    date.getDate(),
  ).padStart(2, '0')}`;
}

test('resolveDefaultLogFilePath uses APPDATA on windows', () => {
  const today = localDateKey(new Date());
  const resolved = resolveDefaultLogFilePath({
    platform: 'win32',
    homeDir: 'C:\\Users\\tester',
    appDataDir: 'C:\\Users\\tester\\AppData\\Roaming',
  });

  assert.equal(
    path.normalize(resolved),
    path.normalize(
      path.join('C:\\Users\\tester\\AppData\\Roaming', 'SubMiner', 'logs', `app-${today}.log`),
    ),
  );
});

test('resolveDefaultLogFilePath uses .config on linux', () => {
  const today = localDateKey(new Date());
  const resolved = resolveDefaultLogFilePath({
    platform: 'linux',
    homeDir: '/home/tester',
  });

  assert.equal(
    resolved,
    path.join('/home/tester', '.config', 'SubMiner', 'logs', `app-${today}.log`),
  );
});

test('setLogRotation accepts numeric retention days', () => {
  const previous = process.env.SUBMINER_LOG_ROTATION;
  const today = localDateKey(new Date());
  setLogRotation(14);
  try {
    const resolved = resolveDefaultLogFilePath({
      platform: 'linux',
      homeDir: '/home/tester',
    });

    assert.equal(
      resolved,
      path.join('/home/tester', '.config', 'SubMiner', 'logs', `app-${today}.log`),
    );
    assert.equal(process.env.SUBMINER_LOG_ROTATION, '14');
  } finally {
    if (previous == null) {
      delete process.env.SUBMINER_LOG_ROTATION;
    } else {
      process.env.SUBMINER_LOG_ROTATION = previous;
    }
  }
});
