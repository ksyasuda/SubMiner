import assert from 'node:assert/strict';
import test from 'node:test';
import { nowMs } from './time.js';

test('nowMs returns wall-clock epoch milliseconds', () => {
  assert.ok(nowMs() > 1_600_000_000_000);
});

test('nowMs honors string-backed test clock values', () => {
  const previousNowMs = globalThis.__subminerTestNowMs;
  globalThis.__subminerTestNowMs = '1700000000123.9';

  try {
    assert.equal(nowMs(), 1_700_000_000_123);
  } finally {
    globalThis.__subminerTestNowMs = previousNowMs;
  }
});
