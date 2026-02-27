import test from 'node:test';
import assert from 'node:assert/strict';

import {
  createOverlayContentMeasurementStore,
  sanitizeOverlayContentMeasurement,
} from './overlay-content-measurement';

test('sanitizeOverlayContentMeasurement accepts valid payload with null rect', () => {
  const measurement = sanitizeOverlayContentMeasurement(
    {
      layer: 'visible',
      measuredAtMs: 100,
      viewport: { width: 1920, height: 1080 },
      contentRect: null,
    },
    500,
  );

  assert.deepEqual(measurement, {
    layer: 'visible',
    measuredAtMs: 100,
    viewport: { width: 1920, height: 1080 },
    contentRect: null,
  });
});

test('sanitizeOverlayContentMeasurement rejects invalid ranges', () => {
  const measurement = sanitizeOverlayContentMeasurement(
    {
      layer: 'visible',
      measuredAtMs: 100,
      viewport: { width: 0, height: 1080 },
      contentRect: { x: 0, y: 0, width: 100, height: 20 },
    },
    500,
  );

  assert.equal(measurement, null);
});

test('overlay measurement store keeps latest payload for visible layer', () => {
  const store = createOverlayContentMeasurementStore({
    now: () => 1000,
    warn: () => {
      // noop
    },
  });

  const visible = store.report({
    layer: 'visible',
    measuredAtMs: 900,
    viewport: { width: 1280, height: 720 },
    contentRect: { x: 50, y: 60, width: 400, height: 80 },
  });

  assert.equal(visible?.layer, 'visible');
  assert.equal(store.getLatestByLayer('visible')?.contentRect?.width, 400);
});

test('overlay measurement store rate-limits invalid payload warnings', () => {
  let now = 1_000;
  const warnings: string[] = [];
  const store = createOverlayContentMeasurementStore({
    now: () => now,
    warn: (message) => {
      warnings.push(message);
    },
  });

  store.report({ layer: 'visible' });
  store.report({ layer: 'visible' });
  assert.equal(warnings.length, 0);

  now = 11_000;
  store.report({ layer: 'visible' });
  assert.equal(warnings.length, 1);
  assert.match(warnings[0]!, /Dropped 3 invalid measurement payload/);
});
