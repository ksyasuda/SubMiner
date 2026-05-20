import assert from 'node:assert/strict';
import test from 'node:test';
import { createCurrentMediaTokenizationGate } from './current-media-tokenization-gate';

test('current media tokenization gate waits until current path is marked ready', async () => {
  const gate = createCurrentMediaTokenizationGate();
  gate.updateCurrentMediaPath('/tmp/video-1.mkv');

  let resolved = false;
  const waitPromise = gate.waitUntilReady('/tmp/video-1.mkv').then(() => {
    resolved = true;
  });

  await Promise.resolve();
  assert.equal(resolved, false);

  gate.markReady('/tmp/video-1.mkv');
  await waitPromise;
  assert.equal(resolved, true);
});

test('current media tokenization gate resolves old waiters when media changes', async () => {
  const gate = createCurrentMediaTokenizationGate();
  gate.updateCurrentMediaPath('/tmp/video-1.mkv');

  let resolved = false;
  const waitPromise = gate.waitUntilReady('/tmp/video-1.mkv').then(() => {
    resolved = true;
  });

  gate.updateCurrentMediaPath('/tmp/video-2.mkv');
  await waitPromise;
  assert.equal(resolved, true);
});

test('current media tokenization gate returns immediately for ready media', async () => {
  const gate = createCurrentMediaTokenizationGate();
  gate.updateCurrentMediaPath('/tmp/video-1.mkv');
  gate.markReady('/tmp/video-1.mkv');

  await gate.waitUntilReady('/tmp/video-1.mkv');
});

test('current media tokenization gate treats later media as ready after warmup completes', async () => {
  const gate = createCurrentMediaTokenizationGate();
  gate.updateCurrentMediaPath('/tmp/video-1.mkv');
  gate.markReady('/tmp/video-1.mkv');
  gate.updateCurrentMediaPath('/tmp/video-2.mkv');

  let resolved = false;
  await gate.waitUntilReady('/tmp/video-2.mkv').then(() => {
    resolved = true;
  });

  assert.equal(resolved, true);
});
