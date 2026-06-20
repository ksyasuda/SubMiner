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

    cache.start('https://youtu.be/demo', { mode: 'background' });

    assert.equal(spawnCalls.length, 1);
    assert.equal(spawnCalls[0]?.command, 'yt-dlp');
    assert.ok(spawnCalls[0]?.args.includes('--no-playlist'));
    assert.ok(spawnCalls[0]?.args.includes('--merge-output-format'));
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
