import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import test from 'node:test';

import { createYoutubeMediaCacheService } from './media-cache';

class FakeYtDlpProcess extends EventEmitter {
  killed = false;
  stdout = new EventEmitter();
  stderr = new EventEmitter();

  kill(): boolean {
    this.killed = true;
    return true;
  }
}

type SpawnCall = {
  command: string;
  args: string[];
  options?: { stdio?: Array<'ignore' | 'pipe'> };
};

function makeTempCacheRoot(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-youtube-media-cache-test-'));
}

test('YouTube media cache does nothing in direct mode', async () => {
  const cacheRoot = makeTempCacheRoot();
  const spawnCalls: SpawnCall[] = [];

  try {
    const cache = createYoutubeMediaCacheService({
      cacheRoot,
      getYtDlpCommand: () => 'yt-dlp',
      spawn: (command, args) => {
        spawnCalls.push({ command, args });
        return new FakeYtDlpProcess();
      },
    });

    cache.start('https://youtu.be/demo', { mode: 'direct' });

    assert.deepEqual(spawnCalls, []);
    assert.equal(await cache.getCachedMediaPath('https://youtu.be/demo'), null);
  } finally {
    fs.rmSync(cacheRoot, { recursive: true, force: true });
  }
});

test('YouTube media cache exposes the downloaded file after the background job completes', async () => {
  const cacheRoot = makeTempCacheRoot();
  const spawnedProcesses: FakeYtDlpProcess[] = [];
  const spawnCalls: SpawnCall[] = [];
  const readyEvents: Array<{ url: string; path: string }> = [];
  const startedEvents: Array<{ url: string }> = [];

  try {
    const cache = createYoutubeMediaCacheService({
      cacheRoot,
      getYtDlpCommand: () => 'yt-dlp',
      onDownloadStarted: (event) => {
        startedEvents.push(event);
      },
      onReady: (event) => {
        readyEvents.push(event);
      },
      spawn: (command, args, options) => {
        spawnCalls.push({ command, args, options });
        const proc = new FakeYtDlpProcess();
        spawnedProcesses.push(proc);
        return proc;
      },
    });

    cache.start('https://youtu.be/demo', { mode: 'background' });

    assert.deepEqual(startedEvents, [{ url: 'https://youtu.be/demo' }]);
    assert.equal(spawnCalls.length, 1);
    assert.equal(spawnCalls[0]?.command, 'yt-dlp');
    assert.ok(spawnCalls[0]?.args.includes('--no-playlist'));
    assert.ok(spawnCalls[0]?.args.includes('--force-ipv4'));
    assert.equal(spawnCalls[0]?.args[spawnCalls[0].args.indexOf('--retries') + 1], '5');
    assert.equal(spawnCalls[0]?.args[spawnCalls[0].args.indexOf('--fragment-retries') + 1], '5');
    assert.equal(spawnCalls[0]?.args[spawnCalls[0].args.indexOf('--extractor-retries') + 1], '5');
    assert.ok(spawnCalls[0]?.args.includes('--merge-output-format'));
    assert.equal(
      spawnCalls[0]?.args[spawnCalls[0].args.indexOf('-f') + 1],
      'bestvideo*[height<=720]+bestaudio/best[height<=720]',
    );
    assert.deepEqual(spawnCalls[0]?.options?.stdio, ['ignore', 'ignore', 'ignore']);
    assert.equal(await cache.getCachedMediaPath('https://youtu.be/demo'), null);

    const outputTemplate = spawnCalls[0]?.args[spawnCalls[0].args.indexOf('-o') + 1];
    assert.equal(typeof outputTemplate, 'string');
    const outputDir = path.dirname(outputTemplate!);
    const outputPath = path.join(outputDir, 'media.mkv');
    fs.mkdirSync(outputDir, { recursive: true });
    fs.writeFileSync(outputPath, 'cached media');
    spawnedProcesses[0]?.emit('close', 0);

    await new Promise((resolve) => setImmediate(resolve));

    assert.equal(await cache.getCachedMediaPath('https://youtu.be/demo'), outputPath);
    assert.deepEqual(readyEvents, [{ url: 'https://youtu.be/demo', path: outputPath }]);
  } finally {
    fs.rmSync(cacheRoot, { recursive: true, force: true });
  }
});

test('YouTube media cache clears the active cache when direct mode starts', async () => {
  const cacheRoot = makeTempCacheRoot();
  const spawnedProcesses: FakeYtDlpProcess[] = [];
  const spawnCalls: SpawnCall[] = [];

  try {
    const cache = createYoutubeMediaCacheService({
      cacheRoot,
      getYtDlpCommand: () => 'yt-dlp',
      spawn: (command, args, options) => {
        spawnCalls.push({ command, args, options });
        const proc = new FakeYtDlpProcess();
        spawnedProcesses.push(proc);
        return proc;
      },
    });

    cache.start('https://youtu.be/background', { mode: 'background' });
    const outputTemplate = spawnCalls[0]?.args[spawnCalls[0].args.indexOf('-o') + 1];
    assert.equal(typeof outputTemplate, 'string');
    const outputDir = path.dirname(outputTemplate!);
    const outputPath = path.join(outputDir, 'media.mkv');
    fs.mkdirSync(outputDir, { recursive: true });
    fs.writeFileSync(outputPath, 'cached media');
    spawnedProcesses[0]?.emit('close', 0);

    await new Promise((resolve) => setImmediate(resolve));
    assert.equal(await cache.getActiveCachedMediaPath(), outputPath);

    cache.start('https://youtu.be/direct', { mode: 'direct' });

    assert.equal(await cache.getActiveCachedMediaPath(), null);
    assert.equal(await cache.getCachedMediaPath('https://youtu.be/background'), outputPath);
  } finally {
    fs.rmSync(cacheRoot, { recursive: true, force: true });
  }
});

test('YouTube media cache cancels the active download when direct mode starts', async () => {
  const cacheRoot = makeTempCacheRoot();
  const spawnedProcesses: FakeYtDlpProcess[] = [];
  const spawnCalls: SpawnCall[] = [];
  const readyEvents: Array<{ url: string; path: string }> = [];

  try {
    const cache = createYoutubeMediaCacheService({
      cacheRoot,
      getYtDlpCommand: () => 'yt-dlp',
      onReady: (event) => {
        readyEvents.push(event);
      },
      spawn: (command, args, options) => {
        spawnCalls.push({ command, args, options });
        const proc = new FakeYtDlpProcess();
        spawnedProcesses.push(proc);
        return proc;
      },
    });

    cache.start('https://youtu.be/background', { mode: 'background' });
    const outputTemplate = spawnCalls[0]?.args[spawnCalls[0].args.indexOf('-o') + 1];
    assert.equal(typeof outputTemplate, 'string');
    const outputDir = path.dirname(outputTemplate!);
    const outputPath = path.join(outputDir, 'media.mkv');
    fs.mkdirSync(outputDir, { recursive: true });
    fs.writeFileSync(outputPath, 'cached media');

    cache.start('https://youtu.be/direct', { mode: 'direct' });
    spawnedProcesses[0]?.emit('close', 0);

    await new Promise((resolve) => setImmediate(resolve));

    assert.equal(spawnedProcesses[0]?.killed, true);
    assert.deepEqual(readyEvents, []);
    assert.equal(await cache.getActiveCachedMediaPath(), null);
    assert.equal(await cache.getCachedMediaPath('https://youtu.be/background'), null);
  } finally {
    fs.rmSync(cacheRoot, { recursive: true, force: true });
  }
});

test('YouTube media cache reports failed background downloads', async () => {
  const cacheRoot = makeTempCacheRoot();
  const spawnedProcesses: FakeYtDlpProcess[] = [];
  const failedEvents: Array<{ url: string }> = [];

  try {
    const cache = createYoutubeMediaCacheService({
      cacheRoot,
      getYtDlpCommand: () => 'yt-dlp',
      onFailed: (event) => {
        failedEvents.push(event);
      },
      spawn: () => {
        const proc = new FakeYtDlpProcess();
        spawnedProcesses.push(proc);
        return proc;
      },
    });

    cache.start('https://youtu.be/demo', { mode: 'background' });
    spawnedProcesses[0]?.emit('close', 1);

    await new Promise((resolve) => setImmediate(resolve));

    assert.deepEqual(failedEvents, [{ url: 'https://youtu.be/demo' }]);
    assert.equal(await cache.getCachedMediaPath('https://youtu.be/demo'), null);
  } finally {
    fs.rmSync(cacheRoot, { recursive: true, force: true });
  }
});

test('YouTube media cache only reports a failed download once when error is followed by close', async () => {
  const cacheRoot = makeTempCacheRoot();
  const spawnedProcesses: FakeYtDlpProcess[] = [];
  const failedEvents: Array<{ url: string }> = [];

  try {
    const cache = createYoutubeMediaCacheService({
      cacheRoot,
      getYtDlpCommand: () => 'yt-dlp',
      onFailed: (event) => {
        failedEvents.push(event);
      },
      spawn: () => {
        const proc = new FakeYtDlpProcess();
        spawnedProcesses.push(proc);
        return proc;
      },
    });

    cache.start('https://youtu.be/demo', { mode: 'background' });
    spawnedProcesses[0]?.emit('error', new Error('spawn failed'));
    spawnedProcesses[0]?.emit('close', 1);

    await new Promise((resolve) => setImmediate(resolve));

    assert.deepEqual(failedEvents, [{ url: 'https://youtu.be/demo' }]);
  } finally {
    fs.rmSync(cacheRoot, { recursive: true, force: true });
  }
});

test('YouTube media cache can disable the download height cap', () => {
  const cacheRoot = makeTempCacheRoot();
  const spawnCalls: SpawnCall[] = [];

  try {
    const cache = createYoutubeMediaCacheService({
      cacheRoot,
      getYtDlpCommand: () => 'yt-dlp',
      spawn: (command, args, options) => {
        spawnCalls.push({ command, args, options });
        return new FakeYtDlpProcess();
      },
    });

    cache.start('https://youtu.be/demo', { mode: 'background', maxHeight: 0 });

    assert.equal(spawnCalls.length, 1);
    assert.equal(
      spawnCalls[0]?.args[spawnCalls[0].args.indexOf('-f') + 1],
      'bestvideo*+bestaudio/best',
    );
  } finally {
    fs.rmSync(cacheRoot, { recursive: true, force: true });
  }
});

test('YouTube media cache applies the configured download height cap', () => {
  const cacheRoot = makeTempCacheRoot();
  const spawnCalls: SpawnCall[] = [];

  try {
    const cache = createYoutubeMediaCacheService({
      cacheRoot,
      getYtDlpCommand: () => 'yt-dlp',
      spawn: (command, args, options) => {
        spawnCalls.push({ command, args, options });
        return new FakeYtDlpProcess();
      },
    });

    cache.start('https://youtu.be/demo', { mode: 'background', maxHeight: 480 });

    assert.equal(spawnCalls.length, 1);
    assert.equal(
      spawnCalls[0]?.args[spawnCalls[0].args.indexOf('-f') + 1],
      'bestvideo*[height<=480]+bestaudio/best[height<=480]',
    );
  } finally {
    fs.rmSync(cacheRoot, { recursive: true, force: true });
  }
});

test('YouTube media cache restarts ready sessions when the height cap changes', async () => {
  const cacheRoot = makeTempCacheRoot();
  const spawnedProcesses: FakeYtDlpProcess[] = [];
  const spawnCalls: SpawnCall[] = [];

  try {
    const cache = createYoutubeMediaCacheService({
      cacheRoot,
      getYtDlpCommand: () => 'yt-dlp',
      spawn: (command, args, options) => {
        spawnCalls.push({ command, args, options });
        const proc = new FakeYtDlpProcess();
        spawnedProcesses.push(proc);
        return proc;
      },
    });

    cache.start('https://youtu.be/demo', { mode: 'background', maxHeight: 720 });
    const firstOutputTemplate = spawnCalls[0]?.args[spawnCalls[0].args.indexOf('-o') + 1];
    assert.equal(typeof firstOutputTemplate, 'string');
    const firstOutputDir = path.dirname(firstOutputTemplate!);
    const firstOutputPath = path.join(firstOutputDir, 'media.mkv');
    fs.mkdirSync(firstOutputDir, { recursive: true });
    fs.writeFileSync(firstOutputPath, 'cached media');
    spawnedProcesses[0]?.emit('close', 0);

    await new Promise((resolve) => setImmediate(resolve));
    assert.equal(await cache.getCachedMediaPath('https://youtu.be/demo'), firstOutputPath);

    cache.start('https://youtu.be/demo', { mode: 'background', maxHeight: 480 });

    assert.equal(spawnCalls.length, 2);
    assert.equal(fs.existsSync(firstOutputPath), false);
    assert.equal(await cache.getCachedMediaPath('https://youtu.be/demo'), null);
    assert.equal(
      spawnCalls[1]?.args[spawnCalls[1].args.indexOf('-f') + 1],
      'bestvideo*[height<=480]+bestaudio/best[height<=480]',
    );
  } finally {
    fs.rmSync(cacheRoot, { recursive: true, force: true });
  }
});

test('YouTube media cache restarts running sessions when the height cap changes', () => {
  const cacheRoot = makeTempCacheRoot();
  const spawnedProcesses: FakeYtDlpProcess[] = [];
  const spawnCalls: SpawnCall[] = [];

  try {
    const cache = createYoutubeMediaCacheService({
      cacheRoot,
      getYtDlpCommand: () => 'yt-dlp',
      spawn: (command, args, options) => {
        spawnCalls.push({ command, args, options });
        const proc = new FakeYtDlpProcess();
        spawnedProcesses.push(proc);
        return proc;
      },
    });

    cache.start('https://youtu.be/demo', { mode: 'background', maxHeight: 720 });
    cache.start('https://youtu.be/demo', { mode: 'background', maxHeight: 480 });

    assert.equal(spawnedProcesses[0]?.killed, true);
    assert.equal(spawnCalls.length, 2);
    assert.equal(
      spawnCalls[1]?.args[spawnCalls[1].args.indexOf('-f') + 1],
      'bestvideo*[height<=480]+bestaudio/best[height<=480]',
    );
  } finally {
    fs.rmSync(cacheRoot, { recursive: true, force: true });
  }
});

test('YouTube media cache removes stale files from previous runs on startup', () => {
  const cacheRoot = makeTempCacheRoot();
  const staleDir = path.join(cacheRoot, 'stale-session');
  const spawnCalls: SpawnCall[] = [];

  try {
    fs.mkdirSync(staleDir, { recursive: true });
    fs.writeFileSync(path.join(staleDir, 'media.mkv'), 'stale cached media');

    createYoutubeMediaCacheService({
      cacheRoot,
      getYtDlpCommand: () => 'yt-dlp',
      spawn: (command, args) => {
        spawnCalls.push({ command, args });
        return new FakeYtDlpProcess();
      },
    });

    assert.equal(fs.existsSync(staleDir), false);
    assert.deepEqual(spawnCalls, []);
  } finally {
    fs.rmSync(cacheRoot, { recursive: true, force: true });
  }
});

test('YouTube media cache restarts when a ready cached file was deleted externally', async () => {
  const cacheRoot = makeTempCacheRoot();
  const spawnedProcesses: FakeYtDlpProcess[] = [];
  const spawnCalls: Array<{ command: string; args: string[] }> = [];

  try {
    const cache = createYoutubeMediaCacheService({
      cacheRoot,
      getYtDlpCommand: () => 'yt-dlp',
      spawn: (command, args) => {
        spawnCalls.push({ command, args });
        const proc = new FakeYtDlpProcess();
        spawnedProcesses.push(proc);
        return proc;
      },
    });

    cache.start('https://youtu.be/demo', { mode: 'background' });
    const outputTemplate = spawnCalls[0]?.args[spawnCalls[0].args.indexOf('-o') + 1];
    assert.equal(typeof outputTemplate, 'string');
    const outputDir = path.dirname(outputTemplate!);
    const outputPath = path.join(outputDir, 'media.mkv');
    fs.mkdirSync(outputDir, { recursive: true });
    fs.writeFileSync(outputPath, 'cached media');
    spawnedProcesses[0]?.emit('close', 0);

    await new Promise((resolve) => setImmediate(resolve));
    assert.equal(await cache.getCachedMediaPath('https://youtu.be/demo'), outputPath);

    fs.rmSync(outputPath);
    cache.start('https://youtu.be/demo', { mode: 'background' });

    assert.equal(spawnCalls.length, 2);
    assert.equal(await cache.getCachedMediaPath('https://youtu.be/demo'), null);
  } finally {
    fs.rmSync(cacheRoot, { recursive: true, force: true });
  }
});

test('YouTube media cache removes stale disk siblings before starting a new cache', () => {
  const cacheRoot = makeTempCacheRoot();
  const staleDir = path.join(cacheRoot, 'stale-sibling');
  const spawnCalls: Array<{ command: string; args: string[] }> = [];

  try {
    const cache = createYoutubeMediaCacheService({
      cacheRoot,
      getYtDlpCommand: () => 'yt-dlp',
      spawn: (command, args) => {
        spawnCalls.push({ command, args });
        return new FakeYtDlpProcess();
      },
    });
    fs.mkdirSync(staleDir, { recursive: true });
    fs.writeFileSync(path.join(staleDir, 'media.mkv'), 'stale cached media');

    cache.start('https://youtu.be/demo', { mode: 'background' });

    assert.equal(fs.existsSync(staleDir), false);
    assert.equal(spawnCalls.length, 1);
  } finally {
    fs.rmSync(cacheRoot, { recursive: true, force: true });
  }
});

test('YouTube media cache drops old sessions when a new background cache starts', async () => {
  const cacheRoot = makeTempCacheRoot();
  const spawnedProcesses: FakeYtDlpProcess[] = [];
  const spawnCalls: Array<{ command: string; args: string[] }> = [];

  try {
    const cache = createYoutubeMediaCacheService({
      cacheRoot,
      getYtDlpCommand: () => 'yt-dlp',
      spawn: (command, args) => {
        spawnCalls.push({ command, args });
        const proc = new FakeYtDlpProcess();
        spawnedProcesses.push(proc);
        return proc;
      },
    });

    cache.start('https://youtu.be/first', { mode: 'background' });
    const firstOutputTemplate = spawnCalls[0]?.args[spawnCalls[0].args.indexOf('-o') + 1];
    assert.equal(typeof firstOutputTemplate, 'string');
    const firstOutputDir = path.dirname(firstOutputTemplate!);
    fs.mkdirSync(firstOutputDir, { recursive: true });
    fs.writeFileSync(path.join(firstOutputDir, 'media.mkv'), 'cached media');

    cache.start('https://youtu.be/second', { mode: 'background' });

    assert.equal(spawnedProcesses[0]?.killed, true);
    assert.equal(fs.existsSync(firstOutputDir), false);
    assert.equal(await cache.getCachedMediaPath('https://youtu.be/first'), null);
    assert.equal(spawnCalls.length, 2);
  } finally {
    fs.rmSync(cacheRoot, { recursive: true, force: true });
  }
});

test('YouTube media cache cleanup kills downloads and removes temp files', async () => {
  const cacheRoot = makeTempCacheRoot();
  const spawnedProcesses: FakeYtDlpProcess[] = [];
  const spawnCalls: Array<{ command: string; args: string[] }> = [];

  try {
    const cache = createYoutubeMediaCacheService({
      cacheRoot,
      getYtDlpCommand: () => 'yt-dlp',
      spawn: (command, args) => {
        spawnCalls.push({ command, args });
        const proc = new FakeYtDlpProcess();
        spawnedProcesses.push(proc);
        return proc;
      },
    });

    cache.start('https://youtu.be/demo', { mode: 'background' });
    const outputTemplate = spawnCalls[0]?.args[spawnCalls[0].args.indexOf('-o') + 1];
    assert.equal(typeof outputTemplate, 'string');
    const outputDir = path.dirname(outputTemplate!);
    fs.mkdirSync(outputDir, { recursive: true });
    fs.writeFileSync(path.join(outputDir, 'media.mkv'), 'cached media');

    cache.cleanup();

    assert.equal(spawnedProcesses[0]?.killed, true);
    assert.equal(fs.existsSync(outputDir), false);
    assert.equal(await cache.getCachedMediaPath('https://youtu.be/demo'), null);
  } finally {
    fs.rmSync(cacheRoot, { recursive: true, force: true });
  }
});
