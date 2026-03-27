import assert from 'node:assert/strict';
import test from 'node:test';
import { createAnilistRateLimiter } from './rate-limiter';

function createTimerHarness() {
  let now = 1_000;
  const waits: number[] = [];
  const originalNow = Date.now;
  const originalSetTimeout = globalThis.setTimeout;

  Date.now = () => now;
  globalThis.setTimeout = ((handler: TimerHandler, timeout?: number) => {
    const waitMs = Number(timeout ?? 0);
    waits.push(waitMs);
    now += waitMs;
    if (typeof handler === 'function') {
      handler();
    }
    return 0 as unknown as ReturnType<typeof setTimeout>;
  }) as unknown as typeof setTimeout;

  return {
    waits,
    advance(ms: number): void {
      now += ms;
    },
    restore(): void {
      Date.now = originalNow;
      globalThis.setTimeout = originalSetTimeout;
    },
  };
}

test('createAnilistRateLimiter waits for the rolling window when capacity is exhausted', async () => {
  const timers = createTimerHarness();
  const limiter = createAnilistRateLimiter(2);

  try {
    await limiter.acquire();
    await limiter.acquire();
    timers.advance(1);
    await limiter.acquire();

    assert.equal(timers.waits.length, 1);
    assert.equal(timers.waits[0], 60_099);
  } finally {
    timers.restore();
  }
});

test('createAnilistRateLimiter pauses until the response reset time', async () => {
  const timers = createTimerHarness();
  const limiter = createAnilistRateLimiter();

  try {
    limiter.recordResponse(
      new Headers({
        'x-ratelimit-remaining': '4',
        'x-ratelimit-reset': '10',
      }),
    );

    await limiter.acquire();

    assert.deepEqual(timers.waits, [9_000]);
  } finally {
    timers.restore();
  }
});

test('createAnilistRateLimiter honors retry-after headers', async () => {
  const timers = createTimerHarness();
  const limiter = createAnilistRateLimiter();

  try {
    limiter.recordResponse(
      new Headers({
        'retry-after': '3',
      }),
    );

    await limiter.acquire();

    assert.deepEqual(timers.waits, [3_000]);
  } finally {
    timers.restore();
  }
});
