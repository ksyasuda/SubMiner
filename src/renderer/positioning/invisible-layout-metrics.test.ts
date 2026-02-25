import { strict as assert } from 'node:assert';
import { afterEach, test } from 'node:test';
import type { MpvSubtitleRenderMetrics } from '../../types';
import {
  applyPlatformFontCompensation,
  calculateOsdScale,
  calculateSubtitleMetrics,
} from './invisible-layout-metrics';

const BASE_METRICS: MpvSubtitleRenderMetrics = {
  subPos: 100,
  subFontSize: 40,
  subScale: 1,
  subMarginY: 34,
  subMarginX: 19,
  subFont: 'sans-serif',
  subSpacing: 0,
  subBold: false,
  subItalic: false,
  subBorderSize: 2,
  subShadowOffset: 0,
  subAssOverride: 'yes',
  subScaleByWindow: false,
  subUseMargins: true,
  osdHeight: 720,
  osdDimensions: {
    w: 1920,
    h: 1080,
    ml: 100,
    mr: 100,
    mt: 80,
    mb: 60,
  },
};

const originalWindow = globalThis.window;

function setWindowDimensions(width: number, height: number, devicePixelRatio: number): void {
  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      innerWidth: width,
      innerHeight: height,
      devicePixelRatio,
    },
  });
}

afterEach(() => {
  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: originalWindow,
  });
});

test('calculateSubtitleMetrics uses video insets for scale-by-video even when subUseMargins is true', () => {
  setWindowDimensions(1920, 1080, 1);

  const ctx = {
    platform: {
      isMacOSPlatform: false,
      isLinuxPlatform: false,
    },
  } as const;

  const result = calculateSubtitleMetrics(ctx as never, BASE_METRICS);

  const expectedPxPerScaledPixel = (1080 - 80 - 60) / 720;
  assert.equal(result.pxPerScaledPixel, expectedPxPerScaledPixel);
  assert.equal(result.effectiveFontSize, BASE_METRICS.subFontSize * expectedPxPerScaledPixel);
});

test('calculateSubtitleMetrics keeps osd insets for positioning even when subUseMargins is true', () => {
  setWindowDimensions(1920, 1080, 1);

  const ctx = {
    platform: {
      isMacOSPlatform: false,
      isLinuxPlatform: false,
    },
  } as const;

  const result = calculateSubtitleMetrics(ctx as never, BASE_METRICS);

  assert.equal(result.leftInset, 100);
  assert.equal(result.rightInset, 100);
  assert.equal(result.topInset, 80);
  assert.equal(result.bottomInset, 60);
  assert.equal(result.horizontalAvailable, 1720);
});

test('applyPlatformFontCompensation applies calibrated macOS factor', () => {
  assert.equal(applyPlatformFontCompensation(100, true), 82);
  assert.equal(applyPlatformFontCompensation(100, false), 100);
});

test('calculateOsdScale snaps near-DPR macOS ratios to devicePixelRatio', () => {
  const metrics = {
    ...BASE_METRICS,
    osdDimensions: {
      w: 3024,
      h: 1701,
      ml: 116,
      mr: 116,
      mt: 28,
      mb: 28,
    },
  };

  const scale = calculateOsdScale(metrics, true, 1728, 972, 2);
  assert.equal(scale, 2);
});
