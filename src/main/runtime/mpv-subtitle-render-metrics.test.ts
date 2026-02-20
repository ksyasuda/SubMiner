import test from 'node:test';
import assert from 'node:assert/strict';
import { createUpdateMpvSubtitleRenderMetricsHandler } from './mpv-subtitle-render-metrics';
import type { MpvSubtitleRenderMetrics } from '../../types';

const BASE_METRICS: MpvSubtitleRenderMetrics = {
  subPos: 100,
  subFontSize: 36,
  subScale: 1,
  subMarginY: 0,
  subMarginX: 0,
  subFont: '',
  subSpacing: 0,
  subBold: false,
  subItalic: false,
  subBorderSize: 0,
  subShadowOffset: 0,
  subAssOverride: 'yes',
  subScaleByWindow: true,
  subUseMargins: true,
  osdHeight: 0,
  osdDimensions: null,
};

test('subtitle render metrics handler no-ops when patch does not change state', () => {
  let metrics = { ...BASE_METRICS };
  let broadcasts = 0;
  const updateMetrics = createUpdateMpvSubtitleRenderMetricsHandler({
    getCurrentMetrics: () => metrics,
    setCurrentMetrics: (next) => {
      metrics = next;
    },
    applyPatch: (current) => ({ next: current, changed: false }),
    broadcastMetrics: () => {
      broadcasts += 1;
    },
  });

  updateMetrics({});
  assert.equal(broadcasts, 0);
});

test('subtitle render metrics handler updates and broadcasts when changed', () => {
  let metrics = { ...BASE_METRICS };
  let broadcasts = 0;
  const updateMetrics = createUpdateMpvSubtitleRenderMetricsHandler({
    getCurrentMetrics: () => metrics,
    setCurrentMetrics: (next) => {
      metrics = next;
    },
    applyPatch: (current, patch) => ({
      next: { ...current, ...patch },
      changed: true,
    }),
    broadcastMetrics: () => {
      broadcasts += 1;
    },
  });

  updateMetrics({ subPos: 80 });
  assert.equal(metrics.subPos, 80);
  assert.equal(broadcasts, 1);
});
