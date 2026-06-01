import { strict as assert } from 'node:assert';
import { test } from 'node:test';
import {
  buildFullWindowShapeRect,
  restoreLinuxOverlayWindowShape,
} from './linux-overlay-window-shape';

test('buildFullWindowShapeRect maps current bounds to a full-window shape', () => {
  assert.deepEqual(buildFullWindowShapeRect({ x: 100, y: 50, width: 1919.6, height: 1080.4 }), {
    x: 0,
    y: 0,
    width: 1920,
    height: 1080,
  });
});

test('buildFullWindowShapeRect rejects invalid dimensions', () => {
  assert.equal(buildFullWindowShapeRect({ x: 0, y: 0, width: 0, height: 1080 }), null);
  assert.equal(buildFullWindowShapeRect({ x: 0, y: 0, width: 1920, height: Number.NaN }), null);
});

test('restoreLinuxOverlayWindowShape restores a full drawable shape', () => {
  const calls: unknown[] = [];

  assert.equal(
    restoreLinuxOverlayWindowShape({
      isDestroyed: () => false,
      getBounds: () => ({ x: 760, y: 152, width: 1920, height: 1080 }),
      setShape: (rects) => calls.push(rects),
    }),
    true,
  );
  assert.deepEqual(calls, [[{ x: 0, y: 0, width: 1920, height: 1080 }]]);
});

test('restoreLinuxOverlayWindowShape skips destroyed or unsupported windows', () => {
  assert.equal(
    restoreLinuxOverlayWindowShape({
      isDestroyed: () => true,
      getBounds: () => ({ x: 0, y: 0, width: 1920, height: 1080 }),
      setShape: () => {
        throw new Error('should not shape destroyed windows');
      },
    }),
    false,
  );
  assert.equal(
    restoreLinuxOverlayWindowShape({
      isDestroyed: () => false,
      getBounds: () => ({ x: 0, y: 0, width: 1920, height: 1080 }),
    }),
    false,
  );
});
