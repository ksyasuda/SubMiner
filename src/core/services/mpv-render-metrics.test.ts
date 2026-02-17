import test from "node:test";
import assert from "node:assert/strict";
import { MpvSubtitleRenderMetrics } from "../../types";
import {
  applyMpvSubtitleRenderMetricsPatch,
  DEFAULT_MPV_SUBTITLE_RENDER_METRICS,
} from "./mpv-render-metrics";

const BASE: MpvSubtitleRenderMetrics = {
  ...DEFAULT_MPV_SUBTITLE_RENDER_METRICS,
};

test("applyMpvSubtitleRenderMetricsPatch returns unchanged on empty patch", () => {
  const { next, changed } = applyMpvSubtitleRenderMetricsPatch(BASE, {});
  assert.equal(changed, false);
  assert.deepEqual(next, BASE);
});

test("applyMpvSubtitleRenderMetricsPatch reports changed when patch modifies value", () => {
  const { next, changed } = applyMpvSubtitleRenderMetricsPatch(BASE, {
    subPos: 95,
  });
  assert.equal(changed, true);
  assert.equal(next.subPos, 95);
});
