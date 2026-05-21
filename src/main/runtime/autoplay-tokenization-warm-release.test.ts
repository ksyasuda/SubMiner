import assert from 'node:assert/strict';
import test from 'node:test';
import { createAutoplayTokenizationWarmRelease } from './autoplay-tokenization-warm-release';

function flushMicrotasks(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

test('autoplay tokenization warm release signals immediately when warmups are ready', () => {
  const calls: string[] = [];
  const release = createAutoplayTokenizationWarmRelease({
    isTokenizationWarmupReady: () => true,
    startTokenizationWarmups: async () => {
      calls.push('warmup');
    },
    getCurrentMediaPath: () => '/tmp/video.mkv',
    signalAutoplayReady: () => calls.push('signal'),
    warn: () => {},
  });

  release('/tmp/video.mkv');

  assert.deepEqual(calls, ['signal']);
});

test('autoplay tokenization warm release primes subtitles before waiting for warmups', async () => {
  const calls: string[] = [];
  let resolveWarmup!: () => void;
  const warmup = new Promise<void>((resolve) => {
    resolveWarmup = resolve;
  });
  const release = createAutoplayTokenizationWarmRelease({
    isTokenizationWarmupReady: () => false,
    startTokenizationWarmups: async () => {
      calls.push('warmup');
      await warmup;
    },
    getCurrentMediaPath: () => '/tmp/video.mkv',
    primeCurrentSubtitle: () => {
      calls.push('prime');
    },
    signalAutoplayReady: () => calls.push('signal'),
    warn: () => {},
  });

  release('/tmp/video.mkv');
  await Promise.resolve();
  assert.deepEqual(calls, ['prime', 'warmup']);

  resolveWarmup();
  await warmup;
  await flushMicrotasks();

  assert.deepEqual(calls, ['prime', 'warmup', 'signal']);
});

test('autoplay tokenization warm release waits for subtitle priming before signaling ready media', async () => {
  const calls: string[] = [];
  let resolvePrime!: () => void;
  const prime = new Promise<void>((resolve) => {
    resolvePrime = resolve;
  });
  const release = createAutoplayTokenizationWarmRelease({
    isTokenizationWarmupReady: () => true,
    startTokenizationWarmups: async () => {
      calls.push('warmup');
    },
    getCurrentMediaPath: () => '/tmp/video.mkv',
    primeCurrentSubtitle: () => {
      calls.push('prime');
      return prime;
    },
    signalAutoplayReady: () => calls.push('signal'),
    warn: () => {},
  });

  release('/tmp/video.mkv');
  await Promise.resolve();

  assert.deepEqual(calls, ['prime']);

  resolvePrime();
  await prime;
  await Promise.resolve();

  assert.deepEqual(calls, ['prime', 'signal']);
});

test('autoplay tokenization warm release waits for warmups before signaling current media', async () => {
  const calls: string[] = [];
  let resolveWarmup!: () => void;
  const warmup = new Promise<void>((resolve) => {
    resolveWarmup = resolve;
  });
  const release = createAutoplayTokenizationWarmRelease({
    isTokenizationWarmupReady: () => false,
    startTokenizationWarmups: async () => {
      calls.push('warmup');
      await warmup;
    },
    getCurrentMediaPath: () => '/tmp/video.mkv',
    signalAutoplayReady: () => calls.push('signal'),
    warn: () => {},
  });

  release('/tmp/video.mkv');
  await Promise.resolve();
  assert.deepEqual(calls, ['warmup']);

  resolveWarmup();
  await warmup;
  await Promise.resolve();

  assert.deepEqual(calls, ['warmup', 'signal']);
});

test('autoplay tokenization warm release skips stale media after warmup resolves', async () => {
  const calls: string[] = [];
  let currentMediaPath = '/tmp/video-2.mkv';
  const release = createAutoplayTokenizationWarmRelease({
    isTokenizationWarmupReady: () => false,
    startTokenizationWarmups: async () => {
      calls.push('warmup');
    },
    getCurrentMediaPath: () => currentMediaPath,
    signalAutoplayReady: () => calls.push('signal'),
    warn: () => {},
  });

  release('/tmp/video-1.mkv');
  await Promise.resolve();
  currentMediaPath = '/tmp/video-3.mkv';
  await Promise.resolve();

  assert.deepEqual(calls, ['warmup']);
});

test('autoplay tokenization warm release skips signaling when current media is cleared', () => {
  const calls: string[] = [];
  const release = createAutoplayTokenizationWarmRelease({
    isTokenizationWarmupReady: () => true,
    startTokenizationWarmups: async () => {
      calls.push('warmup');
    },
    getCurrentMediaPath: () => null,
    signalAutoplayReady: () => calls.push('signal'),
    warn: () => {},
  });

  release('/tmp/video.mkv');

  assert.deepEqual(calls, []);
});
