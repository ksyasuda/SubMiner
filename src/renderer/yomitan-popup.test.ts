import test from 'node:test';
import assert from 'node:assert/strict';
import { YOMITAN_LOOKUP_EVENT, registerYomitanLookupListener } from './yomitan-popup.js';

test('registerYomitanLookupListener forwards the SubMiner Yomitan lookup event', () => {
  const target = new EventTarget();
  const calls: string[] = [];

  const dispose = registerYomitanLookupListener(target, () => {
    calls.push('lookup');
  });

  target.dispatchEvent(new CustomEvent(YOMITAN_LOOKUP_EVENT));
  dispose();
  target.dispatchEvent(new CustomEvent(YOMITAN_LOOKUP_EVENT));

  assert.deepEqual(calls, ['lookup']);
});
