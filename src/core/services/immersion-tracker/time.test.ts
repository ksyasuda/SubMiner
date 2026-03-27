import assert from 'node:assert/strict';
import test from 'node:test';
import { nowMs } from './time.js';

test('nowMs returns wall-clock epoch milliseconds', () => {
  assert.ok(nowMs() > 1_600_000_000_000);
});
