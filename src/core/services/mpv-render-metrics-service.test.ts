import test from "node:test";
import assert from "node:assert/strict";
import { MpvSubtitleRenderMetrics } from "../../types";
import { applyMpvSubtitleRenderMetricsPatchService } from "./mpv-render-metrics-service";

const BASE: MpvSubtitleRenderMetrics = {
  subPos: 100,
  subFontSize: 38,
  subScale: 1,
  subMarginY: 34,
  subMarginX: 19,
  subFont: "sans-serif",
  subSpacing: 0,
  subBold: false,
  subItalic: false,
  subBorderSize: 2.5,
  subShadowOffset: 0,
  subAssOverride: "yes",
  subScaleByWindow: true,
  subUseMargins: true,
  osdHeight: 720,
  osdDimensions: null,
};

test("applyMpvSubtitleRenderMetricsPatchService returns unchanged on empty patch", () => {
  const { next, changed } = applyMpvSubtitleRenderMetricsPatchService(BASE, {});
  assert.equal(changed, false);
  assert.deepEqual(next, BASE);
});

test("applyMpvSubtitleRenderMetricsPatchService reports changed when patch modifies value", () => {
  const { next, changed } = applyMpvSubtitleRenderMetricsPatchService(BASE, {
    subPos: 95,
  });
  assert.equal(changed, true);
  assert.equal(next.subPos, 95);
});
