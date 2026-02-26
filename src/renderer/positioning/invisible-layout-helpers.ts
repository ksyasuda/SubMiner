import type { MpvSubtitleRenderMetrics } from '../../types';
import type { RendererContext } from '../context';

const INVISIBLE_MACOS_VERTICAL_NUDGE_PX = 5;
const INVISIBLE_MACOS_LINE_HEIGHT_SINGLE = '1.08';
const INVISIBLE_MACOS_LINE_HEIGHT_MULTI = '1.35';
const INVISIBLE_MACOS_LINE_HEIGHT_MULTI_DENSE = '1.48';

let fontMetricsCanvas: HTMLCanvasElement | null = null;

export function applyContainerBaseLayout(
  ctx: RendererContext,
  params: {
    horizontalAvailable: number;
    leftInset: number;
    marginX: number;
    hAlign: 0 | 1 | 2;
  },
): void {
  const { horizontalAvailable, leftInset, marginX, hAlign } = params;

  ctx.dom.subtitleContainer.style.position = 'absolute';
  ctx.dom.subtitleContainer.style.maxWidth = `${horizontalAvailable}px`;
  ctx.dom.subtitleContainer.style.width = `${horizontalAvailable}px`;
  ctx.dom.subtitleContainer.style.padding = '0';
  ctx.dom.subtitleContainer.style.background = 'transparent';
  ctx.dom.subtitleContainer.style.marginBottom = '0';
  ctx.dom.subtitleContainer.style.pointerEvents = 'none';
  ctx.dom.subtitleContainer.style.left = `${leftInset + marginX}px`;
  ctx.dom.subtitleContainer.style.right = '';
  ctx.dom.subtitleContainer.style.transform = '';
  ctx.dom.subtitleContainer.style.textAlign = '';

  if (hAlign === 0) {
    ctx.dom.subtitleContainer.style.textAlign = 'left';
    ctx.dom.subtitleRoot.style.textAlign = 'left';
  } else if (hAlign === 2) {
    ctx.dom.subtitleContainer.style.textAlign = 'right';
    ctx.dom.subtitleRoot.style.textAlign = 'right';
  } else {
    ctx.dom.subtitleContainer.style.textAlign = 'center';
    ctx.dom.subtitleRoot.style.textAlign = 'center';
  }

  ctx.dom.subtitleRoot.style.display = 'inline-block';
  ctx.dom.subtitleRoot.style.maxWidth = '100%';
  ctx.dom.subtitleRoot.style.pointerEvents = 'auto';
}

export function applyVerticalPosition(
  ctx: RendererContext,
  params: {
    metrics: MpvSubtitleRenderMetrics;
    renderAreaHeight: number;
    topInset: number;
    bottomInset: number;
    marginY: number;
    borderPx: number;
    shadowPx: number;
    measuredDescentPx: number | null;
    vAlign: 0 | 1 | 2;
  },
): void {
  const baselineCompensationPx = resolveBaselineCompensationPx(
    params.measuredDescentPx,
    params.borderPx,
    params.shadowPx,
  );

  if (params.vAlign === 2) {
    ctx.dom.subtitleContainer.style.top = `${Math.max(
      0,
      params.topInset + params.marginY - baselineCompensationPx,
    )}px`;
    ctx.dom.subtitleContainer.style.bottom = '';
    return;
  }

  if (params.vAlign === 1) {
    ctx.dom.subtitleContainer.style.top = '50%';
    ctx.dom.subtitleContainer.style.bottom = '';
    ctx.dom.subtitleContainer.style.transform = 'translateY(-50%)';
    return;
  }

  const subPosMargin = ((100 - params.metrics.subPos) / 100) * params.renderAreaHeight;
  const effectiveMargin = Math.max(params.marginY, subPosMargin);
  const bottomPx = Math.max(0, params.bottomInset + effectiveMargin + baselineCompensationPx);

  ctx.dom.subtitleContainer.style.top = '';
  ctx.dom.subtitleContainer.style.bottom = `${bottomPx}px`;
}

export function resolveBaselineCompensationPx(
  measuredDescentPx: number | null,
  borderPx: number,
  shadowPx: number,
): number {
  const outlineCompensationPx = Math.max(0, borderPx * 2 + shadowPx);
  if (typeof measuredDescentPx === 'number' && Number.isFinite(measuredDescentPx) && measuredDescentPx > 0) {
    return Math.max(0, measuredDescentPx + outlineCompensationPx);
  }

  return Math.max(0, (borderPx + shadowPx) * 5);
}

function resolveFontFamily(rawFont: string): string {
  const strippedFont = rawFont
    .replace(
      /\s+(Regular|Bold|Italic|Light|Medium|Semi\s*Bold|Extra\s*Bold|Extra\s*Light|Thin|Black|Heavy|Demi\s*Bold|Book|Condensed)\s*$/i,
      '',
    )
    .trim();

  return strippedFont !== rawFont
    ? `"${rawFont}", "${strippedFont}", sans-serif`
    : `"${rawFont}", sans-serif`;
}

export function resolveInvisibleLineHeight(lineCount: number, isMacOSPlatform: boolean): string {
  if (!isMacOSPlatform) return 'normal';
  if (lineCount >= 3) return INVISIBLE_MACOS_LINE_HEIGHT_MULTI_DENSE;
  if (lineCount >= 2) return INVISIBLE_MACOS_LINE_HEIGHT_MULTI;
  return INVISIBLE_MACOS_LINE_HEIGHT_SINGLE;
}

function resolveLetterSpacing(
  spacing: number,
  pxPerScaledPixel: number,
): string {
  if (Math.abs(spacing) > 0.0001) {
    return `${spacing * pxPerScaledPixel}px`;
  }

  return '0px';
}

function measureFontDescentPx(ctx: RendererContext): number | null {
  if (typeof document === 'undefined') return null;
  const computedStyle = getComputedStyle(ctx.dom.subtitleRoot);
  const font = computedStyle.font?.trim();
  if (!font) return null;

  if (!fontMetricsCanvas) {
    fontMetricsCanvas = document.createElement('canvas');
  }

  const context = fontMetricsCanvas.getContext('2d');
  if (!context) return null;

  context.font = font;
  const metrics = context.measureText('Hg漢あ');
  if (!Number.isFinite(metrics.actualBoundingBoxDescent) || metrics.actualBoundingBoxDescent <= 0) {
    return null;
  }
  return metrics.actualBoundingBoxDescent;
}

function applyComputedLineHeightCompensation(
  ctx: RendererContext,
  effectiveFontSize: number,
): void {
  const computedLineHeight = parseFloat(getComputedStyle(ctx.dom.subtitleRoot).lineHeight);
  if (!Number.isFinite(computedLineHeight) || computedLineHeight <= effectiveFontSize) {
    return;
  }

  const halfLeading = (computedLineHeight - effectiveFontSize) / 2;
  if (halfLeading <= 0.5) return;

  const currentBottom = parseFloat(ctx.dom.subtitleContainer.style.bottom);
  if (Number.isFinite(currentBottom)) {
    ctx.dom.subtitleContainer.style.bottom = `${Math.max(0, currentBottom - halfLeading)}px`;
  }

  const currentTop = parseFloat(ctx.dom.subtitleContainer.style.top);
  if (Number.isFinite(currentTop)) {
    ctx.dom.subtitleContainer.style.top = `${Math.max(0, currentTop - halfLeading)}px`;
  }
}

function applyMacOSAdjustments(ctx: RendererContext): void {
  const isMacOSPlatform = ctx.platform.isMacOSPlatform;
  if (!isMacOSPlatform) return;

  const currentBottom = parseFloat(ctx.dom.subtitleContainer.style.bottom);
  if (!Number.isFinite(currentBottom)) return;

  ctx.dom.subtitleContainer.style.bottom = `${Math.max(
    0,
    currentBottom + INVISIBLE_MACOS_VERTICAL_NUDGE_PX,
  )}px`;
}

export function applyTypography(
  ctx: RendererContext,
  params: {
    metrics: MpvSubtitleRenderMetrics;
    pxPerScaledPixel: number;
    effectiveFontSize: number;
  },
): void {
  const isMacOSPlatform = ctx.platform.isMacOSPlatform;
  const lineCount = Math.max(1, ctx.state.currentInvisibleSubtitleLineCount);
  const invisibleLineHeight = resolveInvisibleLineHeight(lineCount, isMacOSPlatform);

  ctx.dom.subtitleRoot.style.setProperty('--invisible-sub-line-height', invisibleLineHeight);
  ctx.dom.subtitleRoot.style.setProperty(
    'line-height',
    invisibleLineHeight,
    isMacOSPlatform ? 'important' : '',
  );
  ctx.dom.subtitleRoot.style.fontFamily = resolveFontFamily(params.metrics.subFont);
  ctx.dom.subtitleRoot.style.setProperty(
    'letter-spacing',
    resolveLetterSpacing(params.metrics.subSpacing, params.pxPerScaledPixel),
    isMacOSPlatform ? 'important' : '',
  );
  ctx.dom.subtitleRoot.style.fontKerning = isMacOSPlatform ? 'auto' : 'none';
  ctx.dom.subtitleRoot.style.fontWeight = params.metrics.subBold ? '700' : '400';
  ctx.dom.subtitleRoot.style.fontStyle = params.metrics.subItalic ? 'italic' : 'normal';
  ctx.dom.subtitleRoot.style.transform = '';
  ctx.dom.subtitleRoot.style.transformOrigin = '';
  ctx.state.invisibleMeasuredDescentPx = measureFontDescentPx(ctx);

  applyComputedLineHeightCompensation(ctx, params.effectiveFontSize);
  applyMacOSAdjustments(ctx);
}
