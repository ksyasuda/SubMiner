import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { appendLogLine, pruneLogFiles, resolveDefaultLogFilePath } from './log-files';

test('resolveDefaultLogFilePath uses app prefix by default', () => {
  const now = new Date('2026-03-22T12:00:00.000Z');
  const resolved = resolveDefaultLogFilePath('app', {
    platform: 'linux',
    homeDir: '/home/tester',
    now,
  });

  assert.equal(
    resolved,
    path.join(
      '/home/tester',
      '.config',
      'SubMiner',
      'logs',
      `app-${now.toISOString().slice(0, 10)}.log`,
    ),
  );
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
