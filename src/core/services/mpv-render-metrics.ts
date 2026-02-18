import { MpvSubtitleRenderMetrics } from "../../types";
import { asBoolean, asFiniteNumber, asString } from "../utils/coerce";

export const DEFAULT_MPV_SUBTITLE_RENDER_METRICS: MpvSubtitleRenderMetrics = {
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

export function sanitizeMpvSubtitleRenderMetrics(
  current: MpvSubtitleRenderMetrics,
  patch: Partial<MpvSubtitleRenderMetrics> | null | undefined,
): MpvSubtitleRenderMetrics {
  if (!patch) return current;
  return updateMpvSubtitleRenderMetrics(current, patch);
}

export function updateMpvSubtitleRenderMetrics(
  current: MpvSubtitleRenderMetrics,
  patch: Partial<MpvSubtitleRenderMetrics>,
): MpvSubtitleRenderMetrics {
  const patchOsd = patch.osdDimensions;
  const nextOsdDimensions =
    patchOsd &&
    typeof patchOsd.w === "number" &&
    typeof patchOsd.h === "number" &&
    typeof patchOsd.ml === "number" &&
    typeof patchOsd.mr === "number" &&
    typeof patchOsd.mt === "number" &&
    typeof patchOsd.mb === "number"
      ? {
          w: asFiniteNumber(patchOsd.w, 0, 1, 100000),
          h: asFiniteNumber(patchOsd.h, 0, 1, 100000),
          ml: asFiniteNumber(patchOsd.ml, 0, 0, 100000),
          mr: asFiniteNumber(patchOsd.mr, 0, 0, 100000),
          mt: asFiniteNumber(patchOsd.mt, 0, 0, 100000),
          mb: asFiniteNumber(patchOsd.mb, 0, 0, 100000),
        }
      : patchOsd === null
        ? null
        : current.osdDimensions;

  return {
    subPos: asFiniteNumber(patch.subPos, current.subPos, 0, 150),
    subFontSize: asFiniteNumber(patch.subFontSize, current.subFontSize, 1, 200),
    subScale: asFiniteNumber(patch.subScale, current.subScale, 0.1, 10),
    subMarginY: asFiniteNumber(patch.subMarginY, current.subMarginY, 0, 200),
    subMarginX: asFiniteNumber(patch.subMarginX, current.subMarginX, 0, 200),
    subFont: asString(patch.subFont, current.subFont),
    subSpacing: asFiniteNumber(patch.subSpacing, current.subSpacing, -100, 100),
    subBold: asBoolean(patch.subBold, current.subBold),
    subItalic: asBoolean(patch.subItalic, current.subItalic),
    subBorderSize: asFiniteNumber(
      patch.subBorderSize,
      current.subBorderSize,
      0,
      100,
    ),
    subShadowOffset: asFiniteNumber(
      patch.subShadowOffset,
      current.subShadowOffset,
      0,
      100,
    ),
    subAssOverride: asString(patch.subAssOverride, current.subAssOverride),
    subScaleByWindow: asBoolean(
      patch.subScaleByWindow,
      current.subScaleByWindow,
    ),
    subUseMargins: asBoolean(patch.subUseMargins, current.subUseMargins),
    osdHeight: asFiniteNumber(patch.osdHeight, current.osdHeight, 1, 10000),
    osdDimensions: nextOsdDimensions,
  };
}

export function applyMpvSubtitleRenderMetricsPatch(
  current: MpvSubtitleRenderMetrics,
  patch: Partial<MpvSubtitleRenderMetrics>,
): { next: MpvSubtitleRenderMetrics; changed: boolean } {
  const next = updateMpvSubtitleRenderMetrics(current, patch);
  const changed =
    next.subPos !== current.subPos ||
    next.subFontSize !== current.subFontSize ||
    next.subScale !== current.subScale ||
    next.subMarginY !== current.subMarginY ||
    next.subMarginX !== current.subMarginX ||
    next.subFont !== current.subFont ||
    next.subSpacing !== current.subSpacing ||
    next.subBold !== current.subBold ||
    next.subItalic !== current.subItalic ||
    next.subBorderSize !== current.subBorderSize ||
    next.subShadowOffset !== current.subShadowOffset ||
    next.subAssOverride !== current.subAssOverride ||
    next.subScaleByWindow !== current.subScaleByWindow ||
    next.subUseMargins !== current.subUseMargins ||
    next.osdHeight !== current.osdHeight ||
    JSON.stringify(next.osdDimensions) !==
      JSON.stringify(current.osdDimensions);
  return { next, changed };
}
