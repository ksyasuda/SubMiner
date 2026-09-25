import assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import test from 'node:test';
import {
  buildRemoteMediaWindowArgs,
  RemoteMediaWindowCache,
  REMOTE_MEDIA_WINDOW_MAX_SECONDS,
  type RemoteMediaWindowCacheOptions,
} from './remote-media-window-cache';

const SOURCE = {
  path: 'https://jellyfin.example/Videos/abc/stream?static=true',
  audioStreamIndex: 2,
};

type ExecFileStub = NonNullable<RemoteMediaWindowCacheOptions['execFile']>;

function createStub(options: { fail?: boolean; empty?: boolean; defer?: boolean } = {}) {
  const calls: string[][] = [];
  const pendingCallbacks: Array<() => void> = [];
  const execFile: ExecFileStub = (_file, args, _options, callback) => {
    calls.push([...args]);
    const finish = (): void => {
      const outputPath = args.at(-1);
      assert.ok(outputPath);
      if (options.fail) {
        callback(Object.assign(new Error('boom'), { code: 1 }));
        return;
      }
      if (!options.empty) {
        fs.writeFileSync(outputPath, 'mkv', 'utf8');
      }
      callback(null);
    };
    if (options.defer) {
      pendingCallbacks.push(finish);
    } else {
      queueMicrotask(finish);
    }
  };
  return {
    calls,
    execFile,
    flush: () => {
      for (const finish of pendingCallbacks.splice(0)) finish();
    },
  };
}

async function withCache(
  stubOptions: Parameters<typeof createStub>[0],
  cacheOptions: Omit<RemoteMediaWindowCacheOptions, 'execFile' | 'tempDir'>,
  run: (cache: RemoteMediaWindowCache, stub: ReturnType<typeof createStub>) => Promise<void>,
): Promise<void> {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-media-window-test-'));
  const stub = createStub(stubOptions);
  const cache = new RemoteMediaWindowCache({
    tempDir,
    execFile: stub.execFile,
    idleTtlMs: 0,
    logDebug: () => undefined,
    ...cacheOptions,
  });
  try {
    await run(cache, stub);
  } finally {
    cache.cleanup();
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
}

function argValue(args: string[], flag: string): string | undefined {
  const index = args.indexOf(flag);
  return index === -1 ? undefined : args[index + 1];
}

test('buildRemoteMediaWindowArgs stream-copies the window with source timestamps intact', () => {
  const args = buildRemoteMediaWindowArgs(
    { ...SOURCE, inputOptions: { reconnect: true, headers: { Referer: 'https://a.example/' } } },
    { startTime: 22.75, endTime: 33 },
    '/tmp/window.mkv',
  );

  const inputIndex = args.indexOf('-i');
  assert.equal(args[inputIndex + 1], SOURCE.path);
  assert.ok(args.indexOf('-reconnect') < inputIndex);
  assert.ok(args.indexOf('-headers') < inputIndex);
  assert.equal(argValue(args, '-ss'), '22.75');
  assert.equal(argValue(args, '-t'), '10.25');
  assert.ok(args.indexOf('-t') < inputIndex);
  assert.deepEqual(args.slice(args.indexOf('-map'), args.indexOf('-map') + 4), [
    '-map',
    '0:v:0?',
    '-map',
    '0:2',
  ]);
  assert.equal(argValue(args, '-c'), 'copy');
  assert.ok(args.includes('-copyts'));
  assert.ok(args.includes('-start_at_zero'));
  assert.equal(argValue(args, '-f'), 'matroska');
  assert.equal(args.at(-1), '/tmp/window.mkv');
});

test('buildRemoteMediaWindowArgs keeps every audio stream when none is selected', () => {
  const args = buildRemoteMediaWindowArgs(
    { path: SOURCE.path, audioStreamIndex: null },
    { startTime: 0, endTime: 5 },
    '/tmp/window.mkv',
  );

  assert.equal(args[args.lastIndexOf('-map') + 1], '0:a');
});

test('acquire downloads once and reuses the window for covered ranges', async () => {
  await withCache({}, {}, async (cache, stub) => {
    const window = await cache.acquire(SOURCE, { startTime: 10, endTime: 14 });

    assert.equal(stub.calls.length, 1);
    assert.equal(argValue(stub.calls[0]!, '-ss'), '9.75');
    assert.equal(argValue(stub.calls[0]!, '-t'), '5.25');
    assert.equal(window.startTime, 9.75);
    assert.equal(window.endTime, 15);
    assert.equal(window.audioStreamIndex, 2);
    assert.ok(fs.existsSync(window.path));
    assert.deepEqual(window.media, {
      path: window.path,
      source: 'remote-window',
      singleResolvedStream: true,
      absoluteTimestamps: true,
    });

    assert.equal(await cache.acquire(SOURCE, { startTime: 11, endTime: 15 }), window);
    assert.equal(await cache.lookup(SOURCE, { startTime: 12, endTime: 12 }), window);
    assert.equal(
      await cache.lookup(
        { path: SOURCE.path, audioStreamIndex: null },
        { startTime: 12, endTime: 13 },
      ),
      window,
    );
    assert.equal(stub.calls.length, 1);
  });
});

test('lookup never downloads and misses on other ranges, sources, or audio streams', async () => {
  await withCache({}, {}, async (cache, stub) => {
    assert.equal(await cache.lookup(SOURCE, { startTime: 10, endTime: 14 }), null);
    assert.equal(stub.calls.length, 0);

    await cache.acquire(SOURCE, { startTime: 10, endTime: 14 });
    assert.equal(await cache.lookup(SOURCE, { startTime: 14, endTime: 16 }), null);
    assert.equal(
      await cache.lookup(
        { path: 'https://other.example/stream', audioStreamIndex: 2 },
        {
          startTime: 11,
          endTime: 12,
        },
      ),
      null,
    );
    assert.equal(
      await cache.lookup(
        { path: SOURCE.path, audioStreamIndex: 3 },
        { startTime: 11, endTime: 12 },
      ),
      null,
    );
    assert.equal(stub.calls.length, 1);
  });
});

test('acquire widens to the union of the old window and replaces the old file', async () => {
  await withCache({}, {}, async (cache, stub) => {
    const first = await cache.acquire(SOURCE, { startTime: 10, endTime: 14 });
    const second = await cache.acquire(SOURCE, { startTime: 8, endTime: 12 });

    assert.equal(stub.calls.length, 2);
    assert.equal(argValue(stub.calls[1]!, '-ss'), '7.75');
    assert.equal(second.startTime, 7.75);
    assert.equal(second.endTime, 15);
    assert.notEqual(second.path, first.path);
    assert.equal(fs.existsSync(first.path), false);
    assert.ok(fs.existsSync(second.path));
    assert.equal(cache.currentWindow, second);
  });
});

test('acquire shares an in-flight download between concurrent callers', async () => {
  await withCache({ defer: true }, {}, async (cache, stub) => {
    const first = cache.acquire(SOURCE, { startTime: 10, endTime: 14 });
    await Promise.resolve();
    const second = cache.acquire(SOURCE, { startTime: 11, endTime: 13 });
    const lookup = cache.lookup(SOURCE, { startTime: 12, endTime: 12 });
    await Promise.resolve();
    assert.equal(stub.calls.length, 1);

    stub.flush();
    const [a, b, c] = await Promise.all([first, second, lookup]);
    assert.equal(a, b);
    assert.equal(a, c);
    assert.equal(stub.calls.length, 1);
  });
});

test('acquire rejects on ffmpeg failure, leaves no file, and can retry', async () => {
  await withCache({ fail: true }, {}, async (cache, stub) => {
    await assert.rejects(
      cache.acquire(SOURCE, { startTime: 10, endTime: 14 }),
      /FFmpeg media window failed: boom/,
    );
    assert.equal(cache.currentWindow, null);
    assert.equal(await cache.lookup(SOURCE, { startTime: 10, endTime: 14 }), null);

    await assert.rejects(cache.acquire(SOURCE, { startTime: 10, endTime: 14 }));
    assert.equal(stub.calls.length, 2);
  });
  await withCache({ empty: true }, {}, async (cache) => {
    await assert.rejects(
      cache.acquire(SOURCE, { startTime: 10, endTime: 14 }),
      /exited without creating a media window/,
    );
  });
});

test('acquire refuses invalid and oversized ranges without spawning ffmpeg', async () => {
  await withCache({}, {}, async (cache, stub) => {
    await assert.rejects(cache.acquire(SOURCE, { startTime: 10, endTime: 10 }), /invalid/);
    await assert.rejects(cache.acquire(SOURCE, { startTime: -1, endTime: 10 }), /invalid/);
    await assert.rejects(
      cache.acquire(SOURCE, { startTime: 0, endTime: REMOTE_MEDIA_WINDOW_MAX_SECONDS + 1 }),
      /too long/,
    );
    assert.equal(stub.calls.length, 0);
  });
});

test('the window is deleted after the idle timeout and on cleanup', async () => {
  await withCache({}, { idleTtlMs: 20 }, async (cache) => {
    const window = await cache.acquire(SOURCE, { startTime: 10, endTime: 14 });
    await new Promise((resolve) => setTimeout(resolve, 60));

    assert.equal(cache.currentWindow, null);
    assert.equal(fs.existsSync(window.path), false);

    const again = await cache.acquire(SOURCE, { startTime: 10, endTime: 14 });
    cache.cleanup();
    assert.equal(fs.existsSync(again.path), false);
    assert.equal(fs.existsSync(path.dirname(again.path)), false);
  });
});
