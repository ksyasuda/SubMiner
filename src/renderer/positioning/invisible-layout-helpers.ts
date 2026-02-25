import type { MpvSubtitleRenderMetrics } from '../../types';
import type { RendererContext } from '../context';

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

function resolveLetterSpacing(
  spacing: number,
  pxPerScaledPixel: number,
): string {
  if (Math.abs(spacing) > 0.0001) {
    return `${spacing * pxPerScaledPixel}px`;
  }

  return '0px';
}

function resolveInvisibleLineHeight(isMacOSPlatform: boolean): string {
  return isMacOSPlatform ? '0.96' : '1';
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

export function applyTypography(
  ctx: RendererContext,
  params: {
    metrics: MpvSubtitleRenderMetrics;
    pxPerScaledPixel: number;
    effectiveFontSize: number;
  },
): void {
  const isMacOSPlatform = ctx.platform.isMacOSPlatform;

  ctx.dom.subtitleRoot.style.setProperty(
    'line-height',
    resolveInvisibleLineHeight(isMacOSPlatform),
    'important',
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
}
