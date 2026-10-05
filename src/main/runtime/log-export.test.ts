import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { exportLogsArchive, maskUsernamesInLogText } from './log-export';

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-log-export-'));
}

function cleanupDir(dirPath: string): void {
  fs.rmSync(dirPath, { recursive: true, force: true });
}

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

function writeLog(logsDir: string, name: string, content: string, mtime: string): string {
  const logPath = path.join(logsDir, name);
  fs.writeFileSync(logPath, content, 'utf8');
  const date = new Date(mtime);
  fs.utimesSync(logPath, date, date);
  return logPath;
}

function readStoredZipEntries(zipPath: string): Map<string, Buffer> {
  const archive = fs.readFileSync(zipPath);
  const entries = new Map<string, Buffer>();
  let cursor = 0;

  while (cursor + 4 <= archive.length) {
    const signature = archive.readUInt32LE(cursor);
    if (signature === 0x02014b50 || signature === 0x06054b50) {
      break;
    }
    assert.equal(signature, 0x04034b50);

    const compressedSize = archive.readUInt32LE(cursor + 18);
    const fileNameLength = archive.readUInt16LE(cursor + 26);
    const extraLength = archive.readUInt16LE(cursor + 28);
    const fileNameStart = cursor + 30;
    const dataStart = fileNameStart + fileNameLength + extraLength;
    const fileName = archive
      .subarray(fileNameStart, fileNameStart + fileNameLength)
      .toString('utf8');
    const data = archive.subarray(dataStart, dataStart + compressedSize);
    entries.set(fileName, Buffer.from(data));
    cursor = dataStart + compressedSize;
  }

  return entries;
}

test('maskUsernamesInLogText redacts linux macOS and Windows home paths', () => {
  const masked = maskUsernamesInLogText(
    [
      '/home/kyle/.config/SubMiner',
      '/Users/kyle/Library/Application Support/SubMiner',
      'C:\\Users\\kyle\\AppData\\Roaming\\SubMiner',
      'C:\\\\Users\\\\kyle\\\\AppData\\\\Roaming\\\\SubMiner',
    ].join('\n'),
  );

  assert.match(masked, /\/home\/<user>\/\.config/);
  assert.match(masked, /\/Users\/<user>\/Library/);
  assert.match(masked, /C:\\Users\\<user>\\AppData/);
  assert.match(masked, /C:\\\\Users\\\\<user>\\\\AppData/);
  assert.doesNotMatch(masked, /kyle/);
});

test('maskUsernamesInLogText redacts headers and yt-dlp cookie arguments', () => {
  const masked = maskUsernamesInLogText(
    [
      'Authorization: Bearer ya29.secret-token',
      'Cookie: SID=session-value; HSID=history-value',
      'Set-Cookie: VISITOR_INFO1_LIVE=visitor; Path=/',
      'x-goog-visitor-id: Cgt2aXNpdG9y',
      'warn yt-dlp header Authorization: Bearer inline-token',
      'yt-dlp --cookies /Users/kyle/cookies.txt --cookies-from-browser chrome:Profile 1',
      'yt-dlp --cookies=/tmp/ytdlp-cookies.txt --cookies-from-browser=firefox:default',
    ].join('\n'),
  );

  assert.match(masked, /Authorization: <redacted>/);
  assert.match(masked, /Cookie: <redacted>/);
  assert.match(masked, /Set-Cookie: <redacted>/);
  assert.match(masked, /x-goog-visitor-id: <redacted>/i);
  assert.match(masked, /--cookies <redacted>/);
  assert.match(masked, /--cookies=<redacted>/);
  assert.match(masked, /--cookies-from-browser <redacted>/);
  assert.match(masked, /--cookies-from-browser=<redacted>/);
  assert.doesNotMatch(masked, /ya29/);
  assert.doesNotMatch(masked, /inline-token/);
  assert.doesNotMatch(masked, /session-value/);
  assert.doesNotMatch(masked, /cookies\.txt/);
  assert.doesNotMatch(masked, /firefox:default/);
});

test('exportLogsArchive exports current-day logs and masks usernames', () => {
  const root = makeTempDir();
  const logsDir = path.join(root, 'logs');
  fs.mkdirSync(logsDir, { recursive: true });

  try {
    const currentLog = writeLog(
      logsDir,
      'app-2026-W21.log',
      'opened /home/kyle/video.mkv and C:\\Users\\kyle\\AppData\\Roaming\\SubMiner\n',
      '2026-05-26T12:00:00.000Z',
    );
    writeLog(logsDir, 'launcher-2026-W20.log', 'old /Users/kyle/Library\n', '2026-05-20T12:00:00Z');

    const result = exportLogsArchive({
      logsDir,
      outputDir: root,
      now: new Date('2026-05-26T16:00:00.000Z'),
    });

    assert.equal(result.mode, 'current-day');
    assert.deepEqual(result.exportedFiles, [currentLog]);

    const entries = readStoredZipEntries(result.zipPath);
    assert.deepEqual([...entries.keys()], ['logs/app-2026-W21.log']);
    const content = entries.get('logs/app-2026-W21.log')!.toString('utf8');
    assert.match(content, /\/home\/<user>\/video\.mkv/);
    assert.match(content, /C:\\Users\\<user>\\AppData/);
    assert.doesNotMatch(content, /kyle/);
  } finally {
    cleanupDir(root);
  }
});

test('exportLogsArchive ignores older dated logs when current-day dated logs exist', () => {
  const root = makeTempDir();
  const logsDir = path.join(root, 'logs');
  fs.mkdirSync(logsDir, { recursive: true });

  try {
    const currentLog = writeLog(
      logsDir,
      'app-2026-05-25.log',
      'current day\n',
      '2026-05-25T18:00:00Z',
    );
    writeLog(logsDir, 'app-2026-05-24.log', 'previous day touched today\n', '2026-05-25T18:00:00Z');

    const result = exportLogsArchive({
      logsDir,
      outputDir: root,
      now: new Date('2026-05-25T20:00:00.000Z'),
    });

    assert.equal(result.mode, 'current-day');
    assert.deepEqual(result.exportedFiles, [currentLog]);

    const entries = readStoredZipEntries(result.zipPath);
    assert.deepEqual([...entries.keys()], ['logs/app-2026-05-25.log']);
  } finally {
    cleanupDir(root);
  }
});

test('exportLogsArchive ranks dated fallback logs by local day freshness', () => {
  withTimeZone('America/Los_Angeles', () => {
    const root = makeTempDir();
    const logsDir = path.join(root, 'logs');
    fs.mkdirSync(logsDir, { recursive: true });

    try {
      const datedLog = writeLog(
        logsDir,
        'app-2026-05-25.log',
        'dated local-day log\n',
        '2026-05-25T12:00:00Z',
      );
      writeLog(
        logsDir,
        'app-2026-05-undated.log',
        'undated touched before local midnight\n',
        '2026-05-26T01:00:00Z',
      );

      const result = exportLogsArchive({
        logsDir,
        outputDir: root,
        now: new Date('2026-05-27T16:00:00.000Z'),
      });

      assert.equal(result.mode, 'most-recent');
      assert.deepEqual(result.exportedFiles, [datedLog]);

      const entries = readStoredZipEntries(result.zipPath);
      assert.deepEqual([...entries.keys()], ['logs/app-2026-05-25.log']);
    } finally {
      cleanupDir(root);
    }
  });
});

test('exportLogsArchive falls back to newest log per kind', () => {
  const root = makeTempDir();
  const logsDir = path.join(root, 'logs');
  fs.mkdirSync(logsDir, { recursive: true });

  try {
    writeLog(logsDir, 'app-2026-W18.log', 'older app\n', '2026-05-01T12:00:00Z');
    const appLog = writeLog(logsDir, 'app-2026-W19.log', 'newer app\n', '2026-05-12T12:00:00Z');
    const mpvLog = writeLog(logsDir, 'mpv-2026-W17.log', 'latest mpv\n', '2026-05-10T12:00:00Z');
    const launcherLog = writeLog(
      logsDir,
      'launcher-2026-W16.log',
      'latest launcher\n',
      '2026-05-09T12:00:00Z',
    );

    const result = exportLogsArchive({
      logsDir,
      outputDir: root,
      now: new Date('2026-05-26T16:00:00.000Z'),
    });

    assert.equal(result.mode, 'most-recent');
    assert.deepEqual(result.exportedFiles.sort(), [appLog, launcherLog, mpvLog].sort());

    const entries = readStoredZipEntries(result.zipPath);
    assert.deepEqual([...entries.keys()].sort(), [
      'logs/app-2026-W19.log',
      'logs/launcher-2026-W16.log',
      'logs/mpv-2026-W17.log',
    ]);
  } finally {
    cleanupDir(root);
  }
});

test('exportLogsArchive fails when no logs exist', () => {
  const root = makeTempDir();
  const logsDir = path.join(root, 'logs');
  fs.mkdirSync(logsDir, { recursive: true });

  try {
    assert.throws(
      () => exportLogsArchive({ logsDir, outputDir: root }),
      /No SubMiner log files found/,
    );
  } finally {
    cleanupDir(root);
  }
});
