import assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import test from 'node:test';

import { buildAnimatedImageVideoFilter, MediaGenerator } from './media-generator';

async function withStubbedFfmpeg(
  run: (generator: MediaGenerator, argsPath: string) => Promise<void>,
  options: {
    logDebug?: (message: string) => void;
    now?: () => number;
  } = {},
): Promise<void> {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-media-generator-test-'));
  const binDir = path.join(root, 'bin');
  const tempDir = path.join(root, 'media');
  const argsPath = path.join(root, 'ffmpeg-args.txt');
  fs.mkdirSync(binDir, { recursive: true });
  const ffmpegStubPath = path.join(binDir, 'ffmpeg-stub.cjs');
  const ffmpegPath = path.join(binDir, process.platform === 'win32' ? 'ffmpeg.cmd' : 'ffmpeg');
  fs.writeFileSync(
    ffmpegStubPath,
    [
      "const fs = require('node:fs');",
      'const args = process.argv.slice(2);',
      "if (args[0] === '-hide_banner' && args[1] === '-encoders') {",
      "  console.log(' V..... libaom-av1');",
      '  process.exit(0);',
      '}',
      "fs.writeFileSync(process.env.SUBMINER_TEST_FFMPEG_ARGS, JSON.stringify(args), 'utf8');",
      'const outputPath = args.at(-1);',
      "fs.writeFileSync(outputPath, 'avif', 'utf8');",
    ].join('\n'),
    'utf8',
  );
  const ffmpegStub =
    process.platform === 'win32'
      ? ['@echo off', `"${process.execPath}" "${ffmpegStubPath}" %*`].join('\r\n')
      : ['#!/bin/sh', `exec "${process.execPath}" "${ffmpegStubPath}" "$@"`].join('\n');
  fs.writeFileSync(ffmpegPath, ffmpegStub, 'utf8');
  if (process.platform !== 'win32') {
    fs.chmodSync(ffmpegPath, 0o755);
  }

  const originalPath = process.env.PATH;
  const originalArgsPath = process.env.SUBMINER_TEST_FFMPEG_ARGS;
  process.env.PATH = `${binDir}${path.delimiter}${originalPath ?? ''}`;
  process.env.SUBMINER_TEST_FFMPEG_ARGS = argsPath;
  const generator = new MediaGenerator(tempDir, options);

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
  return JSON.parse(fs.readFileSync(argsPath, 'utf8')) as string[];
}

test('buildAnimatedImageVideoFilter holds lead-in until the next frame after the audio boundary', () => {
  assert.equal(
    buildAnimatedImageVideoFilter({
      fps: 24,
      maxWidth: 640,
      leadingStillDuration: 1.25,
    }),
    'tpad=start_duration=1.2916666666666667:start_mode=clone,fps=24,scale=w=640:h=-2',
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
    assert.equal(args[args.indexOf('-t') + 1], '3.1');
    assert.equal(args[args.indexOf('-vf') + 1], 'fps=10,scale=w=640:h=-2');
  });
});

test('generateAnimatedImage defaults to unpadded source start and holds through the next frame', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAnimatedImage('/video.mp4', 10, 12, undefined, {
      fps: 10,
      maxWidth: 640,
    });

    const args = readFfmpegArgs(argsPath);
    assert.equal(args[args.indexOf('-ss') + 1], '10');
    assert.equal(args[args.indexOf('-t') + 1], '2.1');
    assert.equal(args[args.indexOf('-vf') + 1], 'fps=10,scale=w=640:h=-2');
  });
});

test('generateAnimatedImage rounds fractional source duration through the next frame boundary', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAnimatedImage('/video.mp4', 10, 12.04, undefined, {
      fps: 10,
      maxWidth: 640,
    });

    const args = readFfmpegArgs(argsPath);
    assert.equal(args[args.indexOf('-ss') + 1], '10');
    assert.equal(args[args.indexOf('-t') + 1], '2.1');
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
    assert.equal(args[args.indexOf('-t') + 1], '3.1');
    assert.equal(
      args[args.indexOf('-vf') + 1],
      'tpad=start_duration=1.3:start_mode=clone,fps=10,scale=w=640:h=-2',
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
    assert.equal(args[args.indexOf('-t') + 1], '1.8');
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

test('generateAudio normalizes sentence audio by default', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAudio('/video.mp4', 10, 12);

    const args = readFfmpegArgs(argsPath);
    assert.equal(args[args.indexOf('-af') + 1], 'loudnorm=I=-23:TP=-2:LRA=11');
  });
});

test('generateAudio can preserve raw sentence audio loudness', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAudio('/video.mp4', 10, 12, 0, null, false);

    const args = readFfmpegArgs(argsPath);
    assert.equal(args.includes('-af'), false);
  });
});

test('generateAudio applies mpv volume after loudness normalization', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAudio('/video.mp4', 10, 12, 0, null, true, 0.42);

    const args = readFfmpegArgs(argsPath);
    assert.equal(args[args.indexOf('-af') + 1], 'loudnorm=I=-23:TP=-2:LRA=11,volume=0.42');
  });
});

test('generateAudio applies mpv volume without loudness normalization', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAudio('/video.mp4', 10, 12, 0, null, false, 0.75);

    const args = readFfmpegArgs(argsPath);
    assert.equal(args[args.indexOf('-af') + 1], 'volume=0.75');
  });
});

test('generateAudio omits no-op mpv volume filters', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAudio('/video.mp4', 10, 12, 0, null, false, 1);

    const args = readFfmpegArgs(argsPath);
    assert.equal(args.includes('-af'), false);
  });
});

test('generateAudio preserves a zero numeric mpv volume', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAudio('/video.mp4', 10, 12, 0, null, false, 0);

    const args = readFfmpegArgs(argsPath);
    assert.equal(args[args.indexOf('-af') + 1], 'volume=0');
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

test('generateAudio recreates missing temp directory before invoking ffmpeg', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    const tempDir = (generator as unknown as { tempDir: string }).tempDir;
    fs.rmSync(tempDir, { recursive: true, force: true });

    await generator.generateAudio('/video.mp4', 10, 12);

    const args = readFfmpegArgs(argsPath);
    const outputPath = args.at(-1);
    assert.equal(typeof outputPath, 'string');
    assert.equal(fs.existsSync(path.dirname(outputPath!)), true);
  });
});

test('generateAudio adds remote input options before the ffmpeg input', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAudio(
      {
        path: 'https://rr1---sn.example.googlevideo.com/videoplayback?mime=audio%2Fwebm',
        inputOptions: {
          reconnect: true,
          userAgent: 'Mozilla/5.0',
          headers: {
            Referer: 'https://www.youtube.com/',
            Origin: 'https://www.youtube.com',
          },
        },
      },
      10,
      12,
    );

    const args = readFfmpegArgs(argsPath);
    const inputIndex = args.indexOf('-i');
    assert.ok(inputIndex > 0);
    assert.ok(args.indexOf('-reconnect') > -1);
    assert.ok(args.indexOf('-reconnect') < inputIndex);
    assert.equal(args[args.indexOf('-reconnect') + 1], '1');
    assert.equal(args[args.indexOf('-reconnect_streamed') + 1], '1');
    assert.equal(args[args.indexOf('-reconnect_on_network_error') + 1], '1');
    assert.equal(args[args.indexOf('-reconnect_on_http_error') + 1], '403,5xx');
    assert.equal(args[args.indexOf('-reconnect_delay_max') + 1], '5');
    assert.equal(args[args.indexOf('-user_agent') + 1], 'Mozilla/5.0');
    assert.equal(
      args[args.indexOf('-headers') + 1],
      'Referer: https://www.youtube.com/\r\nOrigin: https://www.youtube.com\r\n',
    );
  });
});

test('generateAudio skips stale audio stream maps for single resolved streams', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAudio(
      {
        path: 'https://rr1---sn.example.googlevideo.com/videoplayback?mime=audio%2Fwebm',
        singleResolvedStream: true,
      },
      10,
      12,
      0,
      22,
    );

    const args = readFfmpegArgs(argsPath);
    assert.equal(args.includes('-map'), false);
  });
});

test('generateAudio keeps explicit audio stream maps for normal media paths', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAudio('/video.mp4', 10, 12, 0, 2);

    const args = readFfmpegArgs(argsPath);
    assert.equal(args[args.indexOf('-map') + 1], '0:2');
  });
});

test('generateAudio debug-logs cached input and completion timing', async () => {
  const logs: string[] = [];
  const times = [1000, 1052];

  await withStubbedFfmpeg(
    async (generator) => {
      await generator.generateAudio(
        {
          path: '/tmp/subminer-youtube-media-cache/abc123/media.mkv',
          source: 'youtube-cache',
        },
        10,
        12,
      );
    },
    {
      logDebug: (message) => logs.push(message),
      now: () => times.shift() ?? 1052,
    },
  );

  assert.match(logs.join('\n'), /\[media-generator\] audio start/);
  assert.match(logs.join('\n'), /source=youtube-cache/);
  assert.match(
    logs.join('\n'),
    /input=local:\/tmp\/subminer-youtube-media-cache\/abc123\/media\.mkv/,
  );
  assert.match(logs.join('\n'), /\[media-generator\] audio complete/);
  assert.match(logs.join('\n'), /elapsedMs=52/);
  assert.match(logs.join('\n'), /bytes=4/);
});

test('generateAudio debug logs sanitize remote inputs', async () => {
  const logs: string[] = [];
  const times = [1000, 1003];

  await withStubbedFfmpeg(
    async (generator) => {
      await generator.generateAudio(
        {
          path: 'https://rr1---sn.example.googlevideo.com/videoplayback?signature=secret&expire=123',
          inputOptions: {
            reconnect: true,
            headers: {
              Referer: 'https://www.youtube.com/watch?v=abc123',
            },
          },
        },
        10,
        12,
      );
    },
    {
      logDebug: (message) => logs.push(message),
      now: () => times.shift() ?? 1003,
    },
  );

  assert.match(logs.join('\n'), /input=remote:rr1---sn\.example\.googlevideo\.com/);
  assert.doesNotMatch(logs.join('\n'), /signature=secret|expire=123|Referer|abc123/);
});
