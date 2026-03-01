import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildLoadSubtitlePositionMainDepsHandler,
  createBuildSaveSubtitlePositionMainDepsHandler,
} from './subtitle-position-main-deps';

test('load subtitle position main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildLoadSubtitlePositionMainDepsHandler({
    loadSubtitlePositionCore: () => ({ x: 1, y: 2 }) as never,
    setSubtitlePosition: () => calls.push('set'),
  })();

  assert.deepEqual(deps.loadSubtitlePositionCore(), { x: 1, y: 2 });
  deps.setSubtitlePosition({ x: 3, y: 4 } as never);
  assert.deepEqual(calls, ['set']);
});

test('save subtitle position main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildSaveSubtitlePositionMainDepsHandler({
    saveSubtitlePositionCore: () => calls.push('persist'),
    setSubtitlePosition: () => calls.push('set'),
  })();

  deps.setSubtitlePosition({ x: 1, y: 2 } as never);
  deps.saveSubtitlePositionCore({ x: 1, y: 2 } as never);
  assert.deepEqual(calls, ['set', 'persist']);
});
