import test from "node:test";
import assert from "node:assert/strict";
import { createMpvIpcClientDepsRuntimeService } from "./mpv-client-deps-runtime-service";

test("createMpvIpcClientDepsRuntimeService returns passthrough dep object", async () => {
  const marker = {
    getResolvedConfig: () => ({ auto_start_overlay: false } as never),
    autoStartOverlay: true,
    setOverlayVisible: () => {},
    shouldBindVisibleOverlayToMpvSubVisibility: () => true,
    isVisibleOverlayVisible: () => false,
    getReconnectTimer: () => null,
    setReconnectTimer: () => {},
    getCurrentSubText: () => "x",
    setCurrentSubText: () => {},
    setCurrentSubAssText: () => {},
    getSubtitleTimingTracker: () => null,
    subtitleWsBroadcast: () => {},
    getOverlayWindowsCount: () => 0,
    tokenizeSubtitle: async () => ({ text: "x", tokens: [], mergedTokens: [] }),
    broadcastToOverlayWindows: () => {},
    updateCurrentMediaPath: () => {},
    updateMpvSubtitleRenderMetrics: () => {},
    getMpvSubtitleRenderMetrics: () => ({
      subPos: 100,
      subFontSize: 40,
      subScale: 1,
      subMarginY: 0,
      subMarginX: 0,
      subFont: "sans",
      subSpacing: 0,
      subBold: false,
      subItalic: false,
      subBorderSize: 0,
      subShadowOffset: 0,
      subAssOverride: "yes",
      subScaleByWindow: true,
      subUseMargins: true,
      osdHeight: 720,
      osdDimensions: null,
    }),
    setPreviousSecondarySubVisibility: () => {},
    showMpvOsd: () => {},
  };

  const deps = createMpvIpcClientDepsRuntimeService(marker);
  assert.equal(deps.autoStartOverlay, true);
  assert.equal(deps.getCurrentSubText(), "x");
  assert.equal(deps.getOverlayWindowsCount(), 0);
  assert.equal(deps.shouldBindVisibleOverlayToMpvSubVisibility(), true);
});
