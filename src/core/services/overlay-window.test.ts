import test from 'node:test';
import assert from 'node:assert/strict';
import { splitOverlayGeometryForSecondaryBar } from './overlay-window-geometry';

test('splitOverlayGeometryForSecondaryBar returns 20/80 top-bottom regions', () => {
  const regions = splitOverlayGeometryForSecondaryBar({
    x: 100,
    y: 50,
    width: 1200,
    height: 900,
  });

  assert.deepEqual(regions.secondary, {
    x: 100,
    y: 50,
    width: 1200,
    height: 180,
  });
  assert.deepEqual(regions.primary, {
    x: 100,
    y: 230,
    width: 1200,
    height: 720,
  });
});

test('splitOverlayGeometryForSecondaryBar keeps positive sizes for tiny heights', () => {
  const regions = splitOverlayGeometryForSecondaryBar({
    x: 0,
    y: 0,
    width: 300,
    height: 1,
  });

  assert.ok(regions.secondary.height >= 1);
  assert.ok(regions.primary.height >= 1);
});
