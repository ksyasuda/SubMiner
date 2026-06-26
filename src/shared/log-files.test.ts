import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  applyLogFileTogglesToEnv,
  appendLogLine,
  isLogFileEnabled,
  pruneLogDirectoryForPath,
  pruneLogFiles,
  resolveDefaultLogFilePath,
} from './log-files';

function withTimeZone<T>(timeZone: string, run: () => T): T {
  const previous = process.env.TZ;
  process.env.TZ = timeZone;
  try {
    return run();
  } finally {
    if (previous === undefined) {
      delete process.env.TZ;
    } else {
      process.env.TZ = previous;
    }
  }
}

test('resolveDefaultLogFilePath uses app prefix by default', () => {
  const now = new Date('2026-03-22T12:00:00.000Z');
  const resolved = resolveDefaultLogFilePath('app', {
    platform: 'linux',
    homeDir: '/home/tester',
    now,
  });

  assert.equal(
    resolved,
    path.join('/home/tester', '.config', 'SubMiner', 'logs', 'app-2026-03-22.log'),
  );
});

test('resolveDefaultLogFilePath uses the local date when UTC has advanced', () => {
  withTimeZone('America/Los_Angeles', () => {
    const now = new Date('2026-05-26T02:00:00.000Z');
    assert.equal(now.toISOString().slice(0, 10), '2026-05-26');

    const resolved = resolveDefaultLogFilePath('app', {
      platform: 'linux',
      homeDir: '/home/tester',
      now,
    });

    assert.equal(
      resolved,
      path.join('/home/tester', '.config', 'SubMiner', 'logs', 'app-2026-05-25.log'),
    );
  });
});

test('resolveDefaultLogFilePath rejects invalid dates', () => {
  assert.throws(
    () =>
      resolveDefaultLogFilePath('app', {
        platform: 'linux',
        homeDir: '/home/tester',
        now: new Date('invalid'),
      }),
    /Invalid time value/,
  );
});

test('resolveDefaultLogFilePath uses daily filenames for mpv logs', () => {
  const now = new Date('2026-03-22T12:00:00.000Z');
  const resolved = resolveDefaultLogFilePath('mpv', {
    platform: 'linux',
    homeDir: '/home/tester',
    now,
  });

  assert.equal(
    resolved,
    path.join('/home/tester', '.config', 'SubMiner', 'logs', 'mpv-2026-03-22.log'),
  );
});

test('log file toggles keep app and launcher enabled while mpv defaults off', () => {
  assert.equal(isLogFileEnabled('app', {}), true);
  assert.equal(isLogFileEnabled('launcher', {}), true);
  assert.equal(isLogFileEnabled('mpv', {}), false);
  assert.equal(isLogFileEnabled('mpv', { SUBMINER_MPV_LOG: '/tmp/mpv.log' }), true);
  assert.equal(
    isLogFileEnabled('mpv', {
      SUBMINER_MPV_LOG: '/tmp/mpv.log',
      SUBMINER_MPV_LOG_ENABLED: 'false',
    }),
    false,
  );
});

test('applyLogFileTogglesToEnv writes log enable env flags', () => {
  const env: NodeJS.ProcessEnv = {};
  applyLogFileTogglesToEnv({ app: false, launcher: true, mpv: true }, env);

  assert.equal(env.SUBMINER_APP_LOG_ENABLED, 'false');
  assert.equal(env.SUBMINER_LAUNCHER_LOG_ENABLED, 'true');
  assert.equal(env.SUBMINER_MPV_LOG_ENABLED, 'true');
});

test('pruneLogFiles removes logs older than retention window', () => {
  const logsDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-log-prune-'));
  const stalePath = path.join(logsDir, 'app-old.log');
  const freshPath = path.join(logsDir, 'app-fresh.log');
  fs.writeFileSync(stalePath, 'stale\n', 'utf8');
  fs.writeFileSync(freshPath, 'fresh\n', 'utf8');
  const now = new Date('2026-03-22T12:00:00.000Z');
  fs.utimesSync(
    stalePath,
    new Date('2026-03-01T12:00:00.000Z'),
    new Date('2026-03-01T12:00:00.000Z'),
  );
  fs.utimesSync(
    freshPath,
    new Date('2026-03-21T12:00:00.000Z'),
    new Date('2026-03-21T12:00:00.000Z'),
  );

  try {
    pruneLogFiles(logsDir, { retentionDays: 7, now });

    assert.equal(fs.existsSync(stalePath), false);
    assert.equal(fs.existsSync(freshPath), true);
  } finally {
    fs.rmSync(logsDir, { recursive: true, force: true });
  }
});

test('appendLogLine trims oversized logs to newest bytes', () => {
  const logsDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-log-trim-'));
  const logPath = path.join(logsDir, 'app.log');

  try {
    appendLogLine(logPath, '012345678901234567890123456789', { maxBytes: 48, retentionDays: 30 });
    appendLogLine(logPath, 'abcdefghijabcdefghijabcdefghij', { maxBytes: 48, retentionDays: 30 });

    const content = fs.readFileSync(logPath, 'utf8');
    assert.match(content, /\[truncated older log content\]/);
    assert.match(content, /abcdefghij/);
    assert.ok(Buffer.byteLength(content) <= 48);
  } finally {
    fs.rmSync(logsDir, { recursive: true, force: true });
  }
});

test('empty log path operations are no-ops', () => {
  const cwd = process.cwd();
  const logsDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-empty-log-path-'));
  const candidate = path.join(logsDir, 'cwd.log');

  try {
    process.chdir(logsDir);
    fs.writeFileSync(candidate, 'keep\n', 'utf8');

    pruneLogDirectoryForPath('', 1);
    appendLogLine('', 'ignored', { retentionDays: 1 });

    assert.equal(fs.readFileSync(candidate, 'utf8'), 'keep\n');
  } finally {
    process.chdir(cwd);
    fs.rmSync(logsDir, { recursive: true, force: true });
  }
});
