import assert from 'node:assert/strict';
import test from 'node:test';
import type { MpvSubtitleRenderMetrics } from '../../types';
import { createBuildUpdateMpvSubtitleRenderMetricsMainDepsHandler } from './mpv-subtitle-render-metrics-main-deps';

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

test('mpv subtitle render metrics main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildUpdateMpvSubtitleRenderMetricsMainDepsHandler({
    getCurrentMetrics: () => BASE_METRICS,
    setCurrentMetrics: () => calls.push('set'),
    applyPatch: (current, patch) => {
      calls.push('apply');
      return { next: { ...current, ...patch }, changed: true };
    },
    broadcastMetrics: () => calls.push('broadcast'),
  })();

  assert.equal(deps.getCurrentMetrics().subPos, 100);
  deps.setCurrentMetrics(BASE_METRICS);
  const patched = deps.applyPatch(BASE_METRICS, { subPos: 90 });
  deps.broadcastMetrics(BASE_METRICS);

  assert.equal(patched.changed, true);
  assert.equal(patched.next.subPos, 90);
  assert.deepEqual(calls, ['set', 'apply', 'broadcast']);
});
