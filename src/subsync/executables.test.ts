import test from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { findExecutable, resolveExecutable, SUBSYNC_EXECUTABLE_NAMES } from './executables';
import { getSubsyncConfig } from './utils';

function withTempBin(run: (dir: string) => void): void {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-executables-'));
  const previousPath = process.env.PATH;
  try {
    process.env.PATH = dir;
    run(dir);
  } finally {
    process.env.PATH = previousPath;
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

function writeExecutable(dir: string, name: string): string {
  const filePath = path.join(dir, name);
  fs.writeFileSync(filePath, '#!/bin/sh\nexit 0\n', { mode: 0o755 });
  return filePath;
}

test('getSubsyncConfig leaves unset tool paths empty instead of guessing /usr/bin', () => {
  const resolved = getSubsyncConfig({ alass_path: '', ffsubsync_path: '', ffmpeg_path: '' });

  assert.equal(resolved.alassPath, '');
  assert.equal(resolved.ffsubsyncPath, '');
  assert.equal(resolved.ffmpegPath, '');
});

test('getSubsyncConfig trims configured paths', () => {
  const resolved = getSubsyncConfig({ alass_path: '  /opt/homebrew/bin/alass-cli  ' });

  assert.equal(resolved.alassPath, '/opt/homebrew/bin/alass-cli');
});

test('resolveExecutable discovers alass-cli on PATH when config is empty', () => {
  withTempBin((dir) => {
    const expected = writeExecutable(dir, 'alass-cli');

    assert.equal(resolveExecutable('', SUBSYNC_EXECUTABLE_NAMES.alass), expected);
    assert.equal(resolveExecutable(undefined, SUBSYNC_EXECUTABLE_NAMES.alass), expected);
  });
});

test('resolveExecutable prefers alass over alass-cli when both exist', () => {
  withTempBin((dir) => {
    const expected = writeExecutable(dir, 'alass');
    writeExecutable(dir, 'alass-cli');

    assert.equal(resolveExecutable('', SUBSYNC_EXECUTABLE_NAMES.alass), expected);
  });
});

test('resolveExecutable honours an explicit path and does not fall back when it is missing', () => {
  withTempBin((dir) => {
    writeExecutable(dir, 'alass');

    assert.equal(resolveExecutable('/nope/alass', SUBSYNC_EXECUTABLE_NAMES.alass), '');
  });
});

test('resolveExecutable treats a bare configured name as a PATH lookup', () => {
  withTempBin((dir) => {
    const expected = writeExecutable(dir, 'ffmpeg');

    assert.equal(resolveExecutable('ffmpeg', SUBSYNC_EXECUTABLE_NAMES.ffmpeg), expected);
  });
});

test('findExecutable returns empty when nothing matches', () => {
  withTempBin(() => {
    assert.equal(findExecutable(['definitely-not-a-real-binary-xyz']), '');
  });
});
