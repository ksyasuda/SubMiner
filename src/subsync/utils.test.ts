import test from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import {
  buildJellyfinEmbeddedSubtitleUrl,
  candidateExecutableNames,
  codecToExtension,
  extensionToFfmpegMuxer,
  findSubsyncExecutable,
  getSubsyncConfig,
  resolveCommandInvocation,
} from './utils';

test('codecToExtension maps stream/web formats to ffmpeg extractable extensions', () => {
  assert.equal(codecToExtension('subrip'), 'srt');
  assert.equal(codecToExtension('webvtt'), 'vtt');
  assert.equal(codecToExtension('vtt'), 'vtt');
  assert.equal(codecToExtension('ttml'), 'ttml');
  assert.equal(extensionToFfmpegMuxer('vtt'), 'webvtt');
  assert.equal(extensionToFfmpegMuxer('srt'), 'srt');
  assert.equal(extensionToFfmpegMuxer('ass'), 'ass');
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

test('Windows candidate discovery excludes extensionless shims and escapes cmd.exe batch arguments', () => {
  assert.deepEqual(candidateExecutableNames('ffsubsync', 'win32', '.COM;.EXE;.BAT;.CMD'), [
    'ffsubsync.COM',
    'ffsubsync.EXE',
    'ffsubsync.BAT',
    'ffsubsync.CMD',
  ]);
  assert.deepEqual(candidateExecutableNames('ffsubsync.exe', 'win32', '.EXE;.CMD'), [
    'ffsubsync.exe',
  ]);

  const invocation = resolveCommandInvocation(
    'C:\\Tools\\ffsubsync.cmd',
    ['C:\\Videos\\Show & (Remix) [1080p].mkv', '-i', 'C:\\Subs\\ep1.srt'],
    'win32',
  );
  assert.equal(invocation.windowsVerbatimArguments, true);
  assert.equal(invocation.args[0], '/d');
  assert.equal(invocation.args[1], '/s');
  assert.equal(invocation.args[2], '/v:off');
  assert.equal(invocation.args[3], '/c');
  assert.ok(invocation.args[4]!.includes('^&'));
  assert.ok(invocation.args[4]!.includes('^('));
  assert.ok(invocation.args[4]!.includes('^)'));
});

test('buildJellyfinEmbeddedSubtitleUrl derives direct-play subtitle endpoints and rejects non-direct URLs', () => {
  assert.equal(
    buildJellyfinEmbeddedSubtitleUrl(
      'http://pve-main:8096/Videos/item-1/stream?static=true&ApiKey=tok123&MediaSourceId=ms-1',
      2,
      'ass',
    ),
    'http://pve-main:8096/Videos/item-1/ms-1/Subtitles/2/Stream.ass?ApiKey=tok123',
  );
  assert.equal(
    buildJellyfinEmbeddedSubtitleUrl(
      'https://media.example.test/jellyfin/Videos/item-2/stream.mkv?Static=True&api_key=tok456&mediaSourceId=ms-2',
      4,
      'ttml',
    ),
    'https://media.example.test/jellyfin/Videos/item-2/ms-2/Subtitles/4/Stream.srt?ApiKey=tok456',
  );
  // Transcoded or non-static streams must not assume container ff-index matches Jellyfin stream index.
  assert.equal(
    buildJellyfinEmbeddedSubtitleUrl(
      'http://pve-main:8096/Videos/item-1/main.m3u8?ApiKey=tok123&MediaSourceId=ms-1',
      2,
      'srt',
    ),
    null,
  );
  assert.equal(
    buildJellyfinEmbeddedSubtitleUrl(
      'http://pve-main:8096/Videos/item-1/stream?ApiKey=tok123&MediaSourceId=ms-1',
      2,
      'srt',
    ),
    null,
  );
  assert.equal(
    buildJellyfinEmbeddedSubtitleUrl('https://cdn.example.test/episode.mkv', 2, 'srt'),
    null,
  );
  // Leading double-slash in pathname must stay anchored to the media origin and never redirect ApiKey.
  const doubleSlashResult = buildJellyfinEmbeddedSubtitleUrl(
    'https://media.example//collector.example/Videos/item-1/stream?static=true&ApiKey=tok789&MediaSourceId=ms-1',
    2,
    'srt',
  );
  assert.ok(doubleSlashResult !== null);
  const parsedDoubleSlash = new URL(doubleSlashResult);
  assert.equal(parsedDoubleSlash.origin, 'https://media.example');
  assert.equal(parsedDoubleSlash.host, 'media.example');
  assert.equal(parsedDoubleSlash.searchParams.get('ApiKey'), 'tok789');
  assert.equal(
    parsedDoubleSlash.pathname,
    '//collector.example/Videos/item-1/ms-1/Subtitles/2/Stream.srt',
  );
});
