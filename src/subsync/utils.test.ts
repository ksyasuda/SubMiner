import test from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { codecToExtension, findSubsyncExecutable, getSubsyncConfig } from './utils';

test('codecToExtension maps stream/web formats to ffmpeg extractable extensions', () => {
  assert.equal(codecToExtension('subrip'), 'srt');
  assert.equal(codecToExtension('webvtt'), 'vtt');
  assert.equal(codecToExtension('vtt'), 'vtt');
  assert.equal(codecToExtension('ttml'), 'ttml');
});

test('codecToExtension returns null for unsupported codecs', () => {
  assert.equal(codecToExtension('unsupported-codec'), null);
});

test('getSubsyncConfig defaults replace to true', () => {
  assert.equal(getSubsyncConfig(undefined).replace, true);
  assert.equal(getSubsyncConfig({}).replace, true);
});

test('getSubsyncConfig respects explicit replace value', () => {
  assert.equal(getSubsyncConfig({ replace: false }).replace, false);
  assert.equal(getSubsyncConfig({ replace: true }).replace, true);
});

test('getSubsyncConfig defaults mode to auto and respects manual mode', () => {
  assert.equal(getSubsyncConfig(undefined).mode, 'auto');
  assert.equal(getSubsyncConfig({}).mode, 'auto');
  assert.equal(getSubsyncConfig({ mode: 'manual' }).mode, 'manual');
});

test('findSubsyncExecutable skips non-executable files on POSIX and resolves bare overrides', () => {
  if (process.platform === 'win32') return;
  const dirA = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-path-a-'));
  const dirB = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-path-b-'));
  const prevPath = process.env.PATH;
  try {
    const nonExec = path.join(dirA, 'my-custom-tool');
    const execFile = path.join(dirB, 'my-custom-tool');
    fs.writeFileSync(nonExec, '#!/bin/sh\nexit 0\n', { mode: 0o644 });
    fs.chmodSync(nonExec, 0o644);
    fs.writeFileSync(execFile, '#!/bin/sh\nexit 0\n', { mode: 0o755 });
    fs.chmodSync(execFile, 0o755);

    process.env.PATH = `${dirA}${path.delimiter}${dirB}`;
    assert.equal(findSubsyncExecutable('my-custom-tool', ['my-custom-tool']), execFile);
    assert.equal(getSubsyncConfig({ ffmpeg_path: 'my-custom-tool' }).ffmpegPath, execFile);
    assert.equal(
      getSubsyncConfig({ ffmpeg_path: 'missing-custom-tool' }).ffmpegPath,
      'missing-custom-tool',
    );
  } finally {
    process.env.PATH = prevPath;
    fs.rmSync(dirA, { recursive: true, force: true });
    fs.rmSync(dirB, { recursive: true, force: true });
  }
});
