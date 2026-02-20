import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createLoadSubtitlePositionHandler,
  createSaveSubtitlePositionHandler,
} from './subtitle-position';

test('createLoadSubtitlePositionHandler stores loaded value', () => {
  let stored: unknown = null;
  const position = { x: 10, y: 20 };
  const load = createLoadSubtitlePositionHandler({
    loadSubtitlePositionCore: () => position as unknown as never,
    setSubtitlePosition: (value) => {
      stored = value;
    },
  });
  const result = load();
  assert.equal(result, position);
  assert.equal(stored, position);
});

test('createSaveSubtitlePositionHandler stores then persists value', () => {
  const calls: string[] = [];
  const position = { x: 5, y: 7 } as unknown as never;
  const save = createSaveSubtitlePositionHandler({
    saveSubtitlePositionCore: () => {
      calls.push('persist');
    },
    setSubtitlePosition: () => {
      calls.push('store');
    },
  });
  save(position);
  assert.deepEqual(calls, ['store', 'persist']);
});
