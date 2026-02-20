import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildGetAnilistMediaGuessRuntimeStateMainDepsHandler,
  createBuildGetCurrentAnilistMediaKeyMainDepsHandler,
  createBuildResetAnilistMediaGuessStateMainDepsHandler,
  createBuildResetAnilistMediaTrackingMainDepsHandler,
  createBuildSetAnilistMediaGuessRuntimeStateMainDepsHandler,
} from './anilist-media-state-main-deps';

test('get current anilist media key main deps builder maps callbacks', () => {
  const deps = createBuildGetCurrentAnilistMediaKeyMainDepsHandler({
    getCurrentMediaPath: () => '/tmp/video.mkv',
  })();
  assert.equal(deps.getCurrentMediaPath(), '/tmp/video.mkv');
});

test('reset anilist media tracking main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildResetAnilistMediaTrackingMainDepsHandler({
    setMediaKey: () => calls.push('key'),
    setMediaDurationSec: () => calls.push('duration'),
    setMediaGuess: () => calls.push('guess'),
    setMediaGuessPromise: () => calls.push('promise'),
    setLastDurationProbeAtMs: () => calls.push('probe'),
  })();
  deps.setMediaKey(null);
  deps.setMediaDurationSec(null);
  deps.setMediaGuess(null);
  deps.setMediaGuessPromise(null);
  deps.setLastDurationProbeAtMs(0);
  assert.deepEqual(calls, ['key', 'duration', 'guess', 'promise', 'probe']);
});

test('get/set anilist media guess runtime state main deps builders map callbacks', () => {
  const getter = createBuildGetAnilistMediaGuessRuntimeStateMainDepsHandler({
    getMediaKey: () => '/tmp/video.mkv',
    getMediaDurationSec: () => 24,
    getMediaGuess: () => ({ title: 'X' }) as never,
    getMediaGuessPromise: () => Promise.resolve(null) as never,
    getLastDurationProbeAtMs: () => 123,
  })();
  assert.equal(getter.getMediaKey(), '/tmp/video.mkv');
  assert.equal(getter.getMediaDurationSec(), 24);
  assert.equal(getter.getLastDurationProbeAtMs(), 123);

  const calls: string[] = [];
  const setter = createBuildSetAnilistMediaGuessRuntimeStateMainDepsHandler({
    setMediaKey: () => calls.push('key'),
    setMediaDurationSec: () => calls.push('duration'),
    setMediaGuess: () => calls.push('guess'),
    setMediaGuessPromise: () => calls.push('promise'),
    setLastDurationProbeAtMs: () => calls.push('probe'),
  })();
  setter.setMediaKey(null);
  setter.setMediaDurationSec(null);
  setter.setMediaGuess(null);
  setter.setMediaGuessPromise(null);
  setter.setLastDurationProbeAtMs(0);
  assert.deepEqual(calls, ['key', 'duration', 'guess', 'promise', 'probe']);
});

test('reset anilist media guess state main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildResetAnilistMediaGuessStateMainDepsHandler({
    setMediaGuess: () => calls.push('guess'),
    setMediaGuessPromise: () => calls.push('promise'),
  })();
  deps.setMediaGuess(null);
  deps.setMediaGuessPromise(null);
  assert.deepEqual(calls, ['guess', 'promise']);
});
