import type { MpvSubtitleRenderMetrics } from "../../types";
import type { RendererContext } from "../context";
import {
  applyContainerBaseLayout,
  applyTypography,
  applyVerticalPosition,
} from "./invisible-layout-helpers.js";
import {
  calculateSubtitleMetrics,
  calculateSubtitlePosition,
} from "./invisible-layout-metrics.js";

export type MpvSubtitleLayoutController = {
  applyInvisibleSubtitleLayoutFromMpvMetrics: (metrics: MpvSubtitleRenderMetrics, source: string) => void;
};

export function createMpvSubtitleLayoutController(
  ctx: RendererContext,
  applySubtitleFontSize: (fontSize: number) => void,
  options: {
    applyInvisibleSubtitleOffsetPosition: () => void;
    updateInvisiblePositionEditHud: () => void;
  },
): MpvSubtitleLayoutController {
  function applyInvisibleSubtitleLayoutFromMpvMetrics(
    metrics: MpvSubtitleRenderMetrics,
    source: string,
  ): void {
    ctx.state.mpvSubtitleRenderMetrics = metrics;

    const geometry = calculateSubtitleMetrics(ctx, metrics);
    const alignment = calculateSubtitlePosition(metrics, geometry.pxPerScaledPixel, 2);

    applySubtitleFontSize(geometry.effectiveFontSize);
    const effectiveBorderSize = metrics.subBorderSize * geometry.pxPerScaledPixel;

    document.documentElement.style.setProperty(
      "--sub-border-size",
      `${effectiveBorderSize}px`,
    );

    applyContainerBaseLayout(ctx, {
      horizontalAvailable: Math.max(
        0,
        geometry.horizontalAvailable - Math.round(geometry.marginX * 2),
      ),
      leftInset: geometry.leftInset,
      marginX: geometry.marginX,
      hAlign: alignment.hAlign,
    });

    applyVerticalPosition(ctx, {
      metrics,
      renderAreaHeight: geometry.renderAreaHeight,
      topInset: geometry.topInset,
      bottomInset: geometry.bottomInset,
      marginY: geometry.marginY,
      effectiveFontSize: geometry.effectiveFontSize,
      vAlign: alignment.vAlign,
    });

    applyTypography(ctx, {
      metrics,
      pxPerScaledPixel: geometry.pxPerScaledPixel,
      effectiveFontSize: geometry.effectiveFontSize,
    });

    ctx.state.invisibleLayoutBaseLeftPx =
      parseFloat(ctx.dom.subtitleContainer.style.left) || 0;

    const parsedBottom = parseFloat(ctx.dom.subtitleContainer.style.bottom);
    ctx.state.invisibleLayoutBaseBottomPx = Number.isFinite(parsedBottom)
      ? parsedBottom
      : null;

    const parsedTop = parseFloat(ctx.dom.subtitleContainer.style.top);
    ctx.state.invisibleLayoutBaseTopPx = Number.isFinite(parsedTop)
      ? parsedTop
      : null;

    options.applyInvisibleSubtitleOffsetPosition();
    options.updateInvisiblePositionEditHud();

    console.log("[invisible-overlay] Applied mpv subtitle render metrics from", source);
  }

  return {
    applyInvisibleSubtitleLayoutFromMpvMetrics,
  };
}
