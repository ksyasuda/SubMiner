import type { MpvSubtitleRenderMetrics } from '../../types';
import type { RendererContext } from '../context';

export type SubtitleAlignment = { hAlign: 0 | 1 | 2; vAlign: 0 | 1 | 2 };

export type SubtitleLayoutGeometry = {
  renderAreaHeight: number;
  renderAreaWidth: number;
  leftInset: number;
  rightInset: number;
  topInset: number;
  bottomInset: number;
  horizontalAvailable: number;
  marginY: number;
  marginX: number;
  pxPerScaledPixel: number;
  effectiveFontSize: number;
};

export function calculateOsdScale(
  metrics: MpvSubtitleRenderMetrics,
  isMacOSPlatform: boolean,
  viewportWidth: number,
  viewportHeight: number,
  devicePixelRatio: number,
): number {
  const dims = metrics.osdDimensions;

  if (!isMacOSPlatform || !dims) {
    return devicePixelRatio;
  }

  const ratios = [dims.w / Math.max(1, viewportWidth), dims.h / Math.max(1, viewportHeight)].filter(
    (value) => Number.isFinite(value) && value > 0,
  );

  const avgRatio =
    ratios.length > 0
      ? ratios.reduce((sum, value) => sum + value, 0) / ratios.length
      : devicePixelRatio;

  const candidates = [1, devicePixelRatio].filter((candidate, index, list) => {
    if (!Number.isFinite(candidate) || candidate <= 0) return false;
    return list.indexOf(candidate) === index;
  });

  const snappedScale = candidates.reduce((best, candidate) => {
    const bestDistance = Math.abs(avgRatio - best);
    const candidateDistance = Math.abs(avgRatio - candidate);
    return candidateDistance < bestDistance ? candidate : best;
  }, candidates[0] ?? 1);

  if (Math.abs(avgRatio - snappedScale) <= 0.35) {
    return snappedScale;
  }

  return avgRatio > 1.25 ? avgRatio : 1;
}

export function calculateSubtitlePosition(
  _metrics: MpvSubtitleRenderMetrics,
  _scale: number,
  alignment: number,
): SubtitleAlignment {
  return {
    hAlign: ((alignment - 1) % 3) as 0 | 1 | 2,
    vAlign: Math.floor((alignment - 1) / 3) as 0 | 1 | 2,
  };
}

function resolveLinePadding(
  metrics: MpvSubtitleRenderMetrics,
  pxPerScaledPixel: number,
): { marginY: number; marginX: number } {
  return {
    marginY: metrics.subMarginY * pxPerScaledPixel,
    marginX: Math.max(0, metrics.subMarginX * pxPerScaledPixel),
  };
}

export function applyPlatformFontCompensation(
  fontSizePx: number,
  isMacOSPlatform: boolean,
): number {
  return isMacOSPlatform ? fontSizePx * 0.82 : fontSizePx;
}

function calculateGeometry(
  metrics: MpvSubtitleRenderMetrics,
  osdToCssScale: number,
): Omit<SubtitleLayoutGeometry, 'marginY' | 'marginX' | 'pxPerScaledPixel' | 'effectiveFontSize'> {
  const dims = metrics.osdDimensions;
  const renderAreaHeight = dims ? dims.h / osdToCssScale : window.innerHeight;
  const renderAreaWidth = dims ? dims.w / osdToCssScale : window.innerWidth;
  const videoLeftInset = dims ? dims.ml / osdToCssScale : 0;
  const videoRightInset = dims ? dims.mr / osdToCssScale : 0;
  const videoTopInset = dims ? dims.mt / osdToCssScale : 0;
  const videoBottomInset = dims ? dims.mb / osdToCssScale : 0;

  // Keep layout anchored to the same drawable video region represented by osd-dimensions.
  const leftInset = videoLeftInset;
  const rightInset = videoRightInset;
  const topInset = videoTopInset;
  const bottomInset = videoBottomInset;
  const horizontalAvailable = Math.max(0, renderAreaWidth - leftInset - rightInset);

  return {
    renderAreaHeight,
    renderAreaWidth,
    leftInset,
    rightInset,
    topInset,
    bottomInset,
    horizontalAvailable,
  };
}

export function calculateSubtitleMetrics(
  ctx: RendererContext,
  metrics: MpvSubtitleRenderMetrics,
): SubtitleLayoutGeometry {
  const osdToCssScale = calculateOsdScale(
    metrics,
    ctx.platform.isMacOSPlatform,
    window.innerWidth,
    window.innerHeight,
    window.devicePixelRatio || 1,
  );
  const geometry = calculateGeometry(metrics, osdToCssScale);
  const rawVideoTopInset = metrics.osdDimensions ? metrics.osdDimensions.mt / osdToCssScale : 0;
  const rawVideoBottomInset = metrics.osdDimensions ? metrics.osdDimensions.mb / osdToCssScale : 0;
  const videoHeight = geometry.renderAreaHeight - rawVideoTopInset - rawVideoBottomInset;
  const scaleRefHeight = metrics.subScaleByWindow ? geometry.renderAreaHeight : videoHeight;
  const pxPerScaledPixel = Math.max(0.1, scaleRefHeight / 720);
  const computedFontSize =
    metrics.subFontSize * metrics.subScale * (ctx.platform.isLinuxPlatform ? 1 : pxPerScaledPixel);
  const effectiveFontSize = applyPlatformFontCompensation(
    computedFontSize,
    ctx.platform.isMacOSPlatform,
  );
  const spacing = resolveLinePadding(metrics, pxPerScaledPixel);

  return {
    ...geometry,
    marginY: spacing.marginY,
    marginX: spacing.marginX,
    pxPerScaledPixel,
    effectiveFontSize,
  };
}
