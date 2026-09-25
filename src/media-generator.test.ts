import assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import test from 'node:test';

import {
  AUDIO_GENERATION_TIMEOUT_MS,
  buildAnimatedImageVideoFilter,
  MediaGenerator,
  type MediaGeneratorOptions,
} from './media-generator';
import { RemoteMediaWindowCache } from './core/services/remote-media-window-cache';

const REMOTE_STREAM_URL = 'https://jellyfin.example/Videos/abc/stream?static=true';

async function withStubbedFfmpeg(
  run: (generator: MediaGenerator, argsPath: string) => Promise<void>,
  options: MediaGeneratorOptions = {},
  stubOptions: {
    skipOutput?: boolean;
  } = {},
): Promise<void> {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-media-generator-test-'));
  const binDir = path.join(root, 'bin');
  const tempDir = path.join(root, 'media');
  const windowsDir = path.join(root, 'windows');
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
      "fs.appendFileSync(process.env.SUBMINER_TEST_FFMPEG_ARGS, JSON.stringify(args) + '\\n', 'utf8');",
      'const outputPath = args.at(-1);',
      "if (process.env.SUBMINER_TEST_FFMPEG_SKIP_OUTPUT !== '1') {",
      "  fs.writeFileSync(outputPath, 'avif', 'utf8');",
      '}',
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
  const originalSkipOutput = process.env.SUBMINER_TEST_FFMPEG_SKIP_OUTPUT;
  process.env.PATH = `${binDir}${path.delimiter}${originalPath ?? ''}`;
  process.env.SUBMINER_TEST_FFMPEG_ARGS = argsPath;
  if (stubOptions.skipOutput) {
    process.env.SUBMINER_TEST_FFMPEG_SKIP_OUTPUT = '1';
  } else {
    delete process.env.SUBMINER_TEST_FFMPEG_SKIP_OUTPUT;
  }
  // Each test gets its own window cache so remote inputs never leak windows between tests.
  const remoteMediaWindows = new RemoteMediaWindowCache({ tempDir: windowsDir, idleTtlMs: 0 });
  const generator = new MediaGenerator(tempDir, {
    remoteMediaWindows,
    ...options,
  });

  try {
    await run(generator, argsPath);
  } finally {
    generator.cleanup();
    remoteMediaWindows.cleanup();
    process.env.PATH = originalPath;
    if (originalArgsPath === undefined) {
      delete process.env.SUBMINER_TEST_FFMPEG_ARGS;
    } else {
      process.env.SUBMINER_TEST_FFMPEG_ARGS = originalArgsPath;
    }
    if (originalSkipOutput === undefined) {
      delete process.env.SUBMINER_TEST_FFMPEG_SKIP_OUTPUT;
    } else {
      process.env.SUBMINER_TEST_FFMPEG_SKIP_OUTPUT = originalSkipOutput;
    }
    fs.rmSync(root, { recursive: true, force: true });
  }
}

function readAllFfmpegArgs(argsPath: string): string[][] {
  return fs
    .readFileSync(argsPath, 'utf8')
    .split('\n')
    .filter((line) => line.trim().length > 0)
    .map((line) => JSON.parse(line) as string[]);
}

/** Arguments of the most recent ffmpeg invocation. */
function readFfmpegArgs(argsPath: string): string[] {
  return readAllFfmpegArgs(argsPath).at(-1) ?? [];
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

test('generateAudio limits amplified mpv volume after applying gain', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAudio('/video.mp4', 10, 12, 0, null, true, 2);

    const args = readFfmpegArgs(argsPath);
    assert.equal(
      args[args.indexOf('-af') + 1],
      'loudnorm=I=-23:TP=-2:LRA=11,volume=2,alimiter=limit=0.891251:level=false',
    );
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
  await withStubbedFfmpeg(
    async (generator, argsPath) => {
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
    },
    { remoteMediaWindows: null },
  );
});

test('generateAudio downloads a remote window once and extracts from it with absolute seeks', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAudio(
      { path: REMOTE_STREAM_URL, inputOptions: { reconnect: true } },
      10,
      12,
      0.5,
      2,
    );

    const calls = readAllFfmpegArgs(argsPath);
    assert.equal(calls.length, 2);
    const [fetchArgs, audioArgs] = calls as [string[], string[]];
    assert.equal(fetchArgs[fetchArgs.indexOf('-i') + 1], REMOTE_STREAM_URL);
    assert.ok(fetchArgs.indexOf('-reconnect') < fetchArgs.indexOf('-i'));
    assert.equal(fetchArgs[fetchArgs.indexOf('-ss') + 1], '9.25');
    assert.equal(fetchArgs[fetchArgs.lastIndexOf('-map') + 1], '0:2');
    assert.ok(fetchArgs.includes('-copyts'));

    const windowPath = audioArgs[audioArgs.indexOf('-i') + 1];
    assert.ok(windowPath?.endsWith('.mkv'));
    assert.notEqual(windowPath, REMOTE_STREAM_URL);
    assert.equal(audioArgs[audioArgs.indexOf('-ss') + 1], '9.5');
    assert.equal(audioArgs[audioArgs.indexOf('-seek_timestamp') + 1], '1');
    assert.ok(audioArgs.indexOf('-seek_timestamp') < audioArgs.indexOf('-i'));
    assert.equal(audioArgs.includes('-reconnect'), false);
    assert.equal(audioArgs.includes('-map'), false);
    assert.equal(audioArgs.includes('-probesize'), false);
    assert.ok(audioArgs.includes('loudnorm=I=-23:TP=-2:LRA=11'));
  });
});

test('generateScreenshot reuses a downloaded window but never downloads one itself', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateScreenshot(REMOTE_STREAM_URL, 11, { format: 'jpg' });
    let calls = readAllFfmpegArgs(argsPath);
    assert.equal(calls.length, 1);
    assert.equal(calls[0]![calls[0]!.indexOf('-i') + 1], REMOTE_STREAM_URL);

    await generator.generateAudio(REMOTE_STREAM_URL, 10, 12);
    await generator.generateScreenshot(REMOTE_STREAM_URL, 11, { format: 'jpg' });
    await generator.generateScreenshot(REMOTE_STREAM_URL, 40, { format: 'jpg' });

    calls = readAllFfmpegArgs(argsPath);
    assert.equal(calls.length, 5);
    const insideWindow = calls[3]!;
    assert.ok(insideWindow[insideWindow.indexOf('-i') + 1]?.endsWith('.mkv'));
    assert.equal(insideWindow[insideWindow.indexOf('-seek_timestamp') + 1], '1');
    const outsideWindow = calls[4]!;
    assert.equal(outsideWindow[outsideWindow.indexOf('-i') + 1], REMOTE_STREAM_URL);
  });
});

test('generateAnimatedImage downloads the clip window before encoding', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAnimatedImage(REMOTE_STREAM_URL, 10, 12, 0, { fps: 10 });

    const calls = readAllFfmpegArgs(argsPath).filter(
      (args) => args[0] !== '-hide_banner' || args[1] !== '-encoders',
    );
    assert.equal(calls.length, 2);
    assert.equal(calls[0]![calls[0]!.indexOf('-i') + 1], REMOTE_STREAM_URL);
    assert.ok(calls[1]![calls[1]!.indexOf('-i') + 1]?.endsWith('.mkv'));
    assert.equal(calls[1]![calls[1]!.indexOf('-seek_timestamp') + 1], '1');
  });
});

test('generateAudio reads the remote source directly when the window download fails', async () => {
  await withStubbedFfmpeg(
    async (generator, argsPath) => {
      await generator.generateAudio(REMOTE_STREAM_URL, 10, 12);

      const args = readFfmpegArgs(argsPath);
      assert.equal(args[args.indexOf('-i') + 1], REMOTE_STREAM_URL);
      assert.equal(args.includes('-seek_timestamp'), false);
    },
    {
      remoteMediaWindows: new RemoteMediaWindowCache({
        execFile: (_file, _args, _options, callback) =>
          queueMicrotask(() => callback(Object.assign(new Error('offline'), { code: 1 }))),
        idleTtlMs: 0,
        logDebug: () => undefined,
      }),
    },
  );
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

    const [fetchArgs, audioArgs] = readAllFfmpegArgs(argsPath) as [string[], string[]];
    assert.equal(fetchArgs[fetchArgs.lastIndexOf('-map') + 1], '0:a');
    assert.equal(audioArgs.includes('-map'), false);
  });
});

test('generateAudio keeps explicit audio stream maps for normal media paths', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAudio('/video.mp4', 10, 12, 0, 2);

    const args = readFfmpegArgs(argsPath);
    assert.equal(args[args.indexOf('-map') + 1], '0:2');
  });
});

test('generateAudio bounds probing when the selected local audio stream is known', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAudio('/video.mkv', 10, 12, 0, 2);

    const args = readFfmpegArgs(argsPath);
    const inputIndex = args.indexOf('-i');
    assert.ok(args.indexOf('-probesize') > -1);
    assert.ok(args.indexOf('-probesize') < inputIndex);
    assert.equal(args[args.indexOf('-probesize') + 1], '32768');
    assert.ok(args.indexOf('-analyzeduration') < inputIndex);
    assert.equal(args[args.indexOf('-analyzeduration') + 1], '0');
  });
});

test('generateAudio retains normal probing for non-Matroska local media', async () => {
  await withStubbedFfmpeg(async (generator, argsPath) => {
    await generator.generateAudio('/video.mp4', 10, 12, 0, 2);

    const args = readFfmpegArgs(argsPath);
    assert.equal(args.includes('-probesize'), false);
    assert.equal(args.includes('-analyzeduration'), false);
  });
});

test('generateAudio retains a two-minute extraction timeout', async () => {
  let observedTimeout: number | undefined;

  await withStubbedFfmpeg(
    async (generator) => {
      await generator.generateAudio('/video.mp4', 10, 12);
    },
    {
      execFile: (_file, args, options, callback) => {
        observedTimeout = options.timeout;
        const outputPath = args.at(-1);
        assert.ok(outputPath);
        fs.writeFileSync(outputPath, 'mp3', 'utf8');
        queueMicrotask(() => callback(null));
      },
    },
  );

  assert.equal(AUDIO_GENERATION_TIMEOUT_MS, 120_000);
  assert.equal(observedTimeout, AUDIO_GENERATION_TIMEOUT_MS);
});

test('generateAudio reports when ffmpeg exits without creating output', async () => {
  await withStubbedFfmpeg(
    async (generator) => {
      await assert.rejects(
        generator.generateAudio('/video.mp4', 10, 12),
        /FFmpeg audio generation failed: FFmpeg exited without creating an output file/,
      );
    },
    {},
    { skipOutput: true },
  );
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
