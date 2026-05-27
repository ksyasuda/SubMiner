import assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import test from 'node:test';

import { buildAnimatedImageVideoFilter, MediaGenerator } from './media-generator';

async function withStubbedFfmpeg(
  run: (generator: MediaGenerator, argsPath: string) => Promise<void>,
): Promise<void> {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-media-generator-test-'));
  const binDir = path.join(root, 'bin');
  const tempDir = path.join(root, 'media');
  const argsPath = path.join(root, 'ffmpeg-args.txt');
  fs.mkdirSync(binDir, { recursive: true });
  const ffmpegPath = path.join(binDir, 'ffmpeg');
  fs.writeFileSync(
    ffmpegPath,
    [
      '#!/bin/sh',
      'if [ "$1" = "-hide_banner" ] && [ "$2" = "-encoders" ]; then',
      '  echo " V..... libaom-av1"',
      '  exit 0',
      'fi',
      'printf "%s\\n" "$@" > "$SUBMINER_TEST_FFMPEG_ARGS"',
      'out=""',
      'for arg in "$@"; do out="$arg"; done',
      'printf avif > "$out"',
    ].join('\n'),
    'utf8',
  );
  fs.chmodSync(ffmpegPath, 0o755);

  const originalPath = process.env.PATH;
  const originalArgsPath = process.env.SUBMINER_TEST_FFMPEG_ARGS;
  process.env.PATH = `${binDir}${path.delimiter}${originalPath ?? ''}`;
  process.env.SUBMINER_TEST_FFMPEG_ARGS = argsPath;
  const generator = new MediaGenerator(tempDir);

  try {
    await run(generator, argsPath);
  } finally {
    generator.cleanup();
    process.env.PATH = originalPath;
    if (originalArgsPath === undefined) {
      delete process.env.SUBMINER_TEST_FFMPEG_ARGS;
    } else {
      process.env.SUBMINER_TEST_FFMPEG_ARGS = originalArgsPath;
    }
    fs.rmSync(root, { recursive: true, force: true });
  }
}

function readFfmpegArgs(argsPath: string): string[] {
  return fs.readFileSync(argsPath, 'utf8').trim().split('\n');
}

test('buildAnimatedImageVideoFilter prepends a cloned first frame when lead-in is provided', () => {
  assert.equal(
    buildAnimatedImageVideoFilter({
      fps: 10,
      maxWidth: 640,
      leadingStillDuration: 1.25,
    }),
    'tpad=start_duration=1.25:start_mode=clone,fps=10,scale=w=640:h=-2',
  );
});

test('generateAnimatedImage includes leading audio padding in the source range', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAnimatedImage('/video.mp4', 10, 12, 0.5, {
      fps: 10,
      maxWidth: 640,
    });

    const args = readFfmpegArgs(argsPath);
    assert.equal(args[args.indexOf('-ss') + 1], '9.5');
    assert.equal(args[args.indexOf('-t') + 1], '3');
    assert.equal(args[args.indexOf('-vf') + 1], 'fps=10,scale=w=640:h=-2');
  });
});

test('generateAnimatedImage defaults to unpadded sentence timing', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAnimatedImage('/video.mp4', 10, 12, undefined, {
      fps: 10,
      maxWidth: 640,
    });

    const args = readFfmpegArgs(argsPath);
    assert.equal(args[args.indexOf('-ss') + 1], '10');
    assert.equal(args[args.indexOf('-t') + 1], '2');
    assert.equal(args[args.indexOf('-vf') + 1], 'fps=10,scale=w=640:h=-2');
  });
});

test('generateAnimatedImage keeps word-audio lead-in separate from audio padding', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAnimatedImage('/video.mp4', 10, 12, 0.5, {
      fps: 10,
      maxWidth: 640,
      leadingStillDuration: 1.25,
    });

    const args = readFfmpegArgs(argsPath);
    assert.equal(args[args.indexOf('-ss') + 1], '9.5');
    assert.equal(args[args.indexOf('-t') + 1], '3');
    assert.equal(
      args[args.indexOf('-vf') + 1],
      'tpad=start_duration=1.25:start_mode=clone,fps=10,scale=w=640:h=-2',
    );
  });
});

test('generateAnimatedImage clips padded source range at the start of media', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAnimatedImage('/video.mp4', 0.2, 1.2, 0.5, {
      fps: 10,
      maxWidth: 640,
    });

    const args = readFfmpegArgs(argsPath);
    assert.equal(args[args.indexOf('-ss') + 1], '0');
    assert.equal(args[args.indexOf('-t') + 1], '1.7');
    assert.equal(args[args.indexOf('-vf') + 1], 'fps=10,scale=w=640:h=-2');
  });
});

test('generateAudio defaults to unpadded sentence timing', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAudio('/video.mp4', 10, 12);

    const args = readFfmpegArgs(argsPath);
    assert.equal(args[args.indexOf('-ss') + 1], '10');
    assert.equal(args[args.indexOf('-t') + 1], '2');
  });
});

test('generateAudio clips leading padding without adding it to trailing duration', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAudio('/video.mp4', 0.2, 1.2, 0.5);

    const args = readFfmpegArgs(argsPath);
    assert.equal(args[args.indexOf('-ss') + 1], '0');
    assert.equal(args[args.indexOf('-t') + 1], '1.7');
  });
});
