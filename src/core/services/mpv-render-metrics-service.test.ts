import test from "node:test";
import assert from "node:assert/strict";
import { MpvSubtitleRenderMetrics } from "../../types";
import {
  applyMpvSubtitleRenderMetricsPatchService,
  DEFAULT_MPV_SUBTITLE_RENDER_METRICS,
} from "./mpv-render-metrics-service";

const BASE: MpvSubtitleRenderMetrics = {
  ...DEFAULT_MPV_SUBTITLE_RENDER_METRICS,
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
