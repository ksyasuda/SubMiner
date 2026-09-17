import assert from 'node:assert/strict';
import test from 'node:test';
import { waitForPlaybackWindow } from './playback-window-ready';

function createClock() {
  let time = 0;
  return {
    now: () => time,
    wait: async (ms: number) => {
      time += ms;
    },
  };
}

test('resolves once the video output is configured and the tracker has the window', async () => {
  const clock = createClock();
  let probes = 0;
  let tracked = false;
  const ready = await waitForPlaybackWindow({
    ...clock,
    isWindowTracked: () => tracked,
    readProperty: async () => {
      probes += 1;
      if (probes === 1) throw new Error('property unavailable');
      if (probes === 3) tracked = true;
      return probes >= 2;
    },
    probeIntervalMs: 100,
  });

  assert.equal(ready, true);
  assert.equal(probes, 3);
  assert.equal(clock.now(), 200);
});

test('a configured video output is enough when no window tracker is running', async () => {
  const ready = await waitForPlaybackWindow({
    ...createClock(),
    isWindowTracked: () => null,
    readProperty: async () => true,
  });

  assert.equal(ready, true);
});

test('gives up after the wall-clock budget when no window appears', async () => {
  const clock = createClock();
  const ready = await waitForPlaybackWindow({
    ...clock,
    isWindowTracked: () => false,
    readProperty: async () => true,
    timeoutMs: 1000,
    probeIntervalMs: 300,
  });

  assert.equal(ready, false);
  assert.equal(clock.now(), 1000);
});
