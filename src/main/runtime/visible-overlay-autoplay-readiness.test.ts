import assert from 'node:assert/strict';
import test from 'node:test';
import { isVisibleOverlayAutoplayTargetReady } from './visible-overlay-autoplay-readiness';
import type { OverlayContentMeasurement } from '../../types';

const visibleMeasurement = (
  measuredAtMs: number,
  rect = { x: 100, y: 800, width: 500, height: 90 },
): OverlayContentMeasurement => ({
  layer: 'visible',
  measuredAtMs,
  viewport: { width: 1920, height: 1080 },
  contentRect: rect,
  interactiveRects: [rect],
});

test('visible overlay autoplay target waits for a fresh interactive subtitle measurement', () => {
  let measurement: OverlayContentMeasurement | null = null;
  const deps = {
    getVisibleOverlayVisible: () => true,
    isOverlayWindowReady: () => true,
    getLatestVisibleMeasurement: () => measurement,
  };
  const signal = {
    mediaPath: '/media/video.mkv',
    payload: { text: '字幕', tokens: null },
    requestedAtMs: 1_000,
  };

  assert.equal(isVisibleOverlayAutoplayTargetReady(deps, signal), false);

  measurement = visibleMeasurement(999);
  assert.equal(isVisibleOverlayAutoplayTargetReady(deps, signal), false);

  measurement = visibleMeasurement(1_000, { x: 100, y: 800, width: 0, height: 90 });
  assert.equal(isVisibleOverlayAutoplayTargetReady(deps, signal), false);

  measurement = visibleMeasurement(1_001);
  assert.equal(isVisibleOverlayAutoplayTargetReady(deps, signal), true);
});

test('visible overlay autoplay target falls back when interactive rects have no area', () => {
  const ready = isVisibleOverlayAutoplayTargetReady(
    {
      getVisibleOverlayVisible: () => true,
      isOverlayWindowReady: () => true,
      getLatestVisibleMeasurement: () => ({
        layer: 'visible',
        measuredAtMs: 2_000,
        viewport: { width: 1920, height: 1080 },
        contentRect: { x: 100, y: 800, width: 500, height: 90 },
        interactiveRects: [{ x: 100, y: 800, width: 0, height: 90 }],
      }),
    },
    {
      mediaPath: '/media/video.mkv',
      payload: { text: '字幕', tokens: null },
      requestedAtMs: 1_000,
    },
  );

  assert.equal(ready, true);
});

test('visible overlay autoplay target accepts synthetic warmup readiness after content-ready', () => {
  const ready = isVisibleOverlayAutoplayTargetReady(
    {
      getVisibleOverlayVisible: () => true,
      isOverlayWindowReady: () => true,
      getLatestVisibleMeasurement: () => null,
    },
    {
      mediaPath: '/media/video.mkv',
      payload: { text: '__warm__', tokens: null },
      requestedAtMs: 1_000,
    },
  );

  assert.equal(ready, true);
});

test('visible overlay autoplay target waits for content-ready before synthetic warmup readiness', () => {
  const ready = isVisibleOverlayAutoplayTargetReady(
    {
      getVisibleOverlayVisible: () => true,
      isOverlayWindowReady: () => false,
      getLatestVisibleMeasurement: () => visibleMeasurement(2_000),
    },
    {
      mediaPath: '/media/video.mkv',
      payload: { text: '__warm__', tokens: null },
      requestedAtMs: 1_000,
    },
  );

  assert.equal(ready, false);
});

test('visible overlay autoplay target rejects empty readiness payloads', () => {
  const ready = isVisibleOverlayAutoplayTargetReady(
    {
      getVisibleOverlayVisible: () => true,
      isOverlayWindowReady: () => true,
      getLatestVisibleMeasurement: () => visibleMeasurement(2_000),
    },
    {
      mediaPath: '/media/video.mkv',
      payload: { text: '', tokens: null },
      requestedAtMs: 1_000,
    },
  );

  assert.equal(ready, false);
});

test('visible overlay autoplay target bypasses measurement when visible overlay is hidden', () => {
  const ready = isVisibleOverlayAutoplayTargetReady(
    {
      getVisibleOverlayVisible: () => false,
      isOverlayWindowReady: () => false,
      getLatestVisibleMeasurement: () => null,
    },
    {
      mediaPath: '/media/video.mkv',
      payload: { text: '__warm__', tokens: null },
      requestedAtMs: 1_000,
    },
  );

  assert.equal(ready, true);
});
