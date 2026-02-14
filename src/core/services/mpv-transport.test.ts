import test from "node:test";
import assert from "node:assert/strict";
import {
  getMpvReconnectDelay,
  scheduleMpvReconnect,
} from "./mpv-transport";

test("getMpvReconnectDelay follows existing reconnect ramp", () => {
  assert.equal(getMpvReconnectDelay(0, true), 1000);
  assert.equal(getMpvReconnectDelay(1, true), 1000);
  assert.equal(getMpvReconnectDelay(2, true), 2000);
  assert.equal(getMpvReconnectDelay(4, true), 5000);
  assert.equal(getMpvReconnectDelay(7, true), 10000);

  assert.equal(getMpvReconnectDelay(0, false), 200);
  assert.equal(getMpvReconnectDelay(2, false), 500);
  assert.equal(getMpvReconnectDelay(4, false), 1000);
  assert.equal(getMpvReconnectDelay(6, false), 2000);
});

test("scheduleMpvReconnect clears existing timer and increments attempt", () => {
  const existing = {} as ReturnType<typeof setTimeout>;
  const cleared: Array<ReturnType<typeof setTimeout> | null> = [];
  const setTimers: Array<ReturnType<typeof setTimeout> | null> = [];
  const calls: Array<{ attempt: number; delay: number }> = [];
  let connected = 0;

  const originalSetTimeout = globalThis.setTimeout;
  const originalClearTimeout = globalThis.clearTimeout;
  (globalThis as any).setTimeout = (handler: () => void, _delay: number) => {
    handler();
    return 1 as unknown as ReturnType<typeof setTimeout>;
  };
  (globalThis as any).clearTimeout = (timer: ReturnType<typeof setTimeout> | null) => {
    cleared.push(timer);
  };

  const nextAttempt = scheduleMpvReconnect({
    attempt: 3,
    hasConnectedOnce: true,
    getReconnectTimer: () => existing,
    setReconnectTimer: (timer) => {
      setTimers.push(timer);
    },
    onReconnectAttempt: (attempt, delay) => {
      calls.push({ attempt, delay });
    },
    connect: () => {
      connected += 1;
    },
  });

  (globalThis as any).setTimeout = originalSetTimeout;
  (globalThis as any).clearTimeout = originalClearTimeout;

  assert.equal(nextAttempt, 4);
  assert.equal(cleared.length, 1);
  assert.equal(cleared[0], existing);
  assert.equal(setTimers.length, 1);
  assert.equal(calls.length, 1);
  assert.equal(calls[0].attempt, 4);
  assert.equal(calls[0].delay, getMpvReconnectDelay(3, true));
  assert.equal(connected, 1);
});
