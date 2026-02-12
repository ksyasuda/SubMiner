import type { MpvSubtitleRenderMetrics, SubtitlePosition } from "../types";
import type { ModalStateReader, RendererContext } from "./context";

const INVISIBLE_MACOS_VERTICAL_NUDGE_PX = 5;
const INVISIBLE_MACOS_LINE_HEIGHT_SINGLE = "0.92";
const INVISIBLE_MACOS_LINE_HEIGHT_MULTI = "1.2";
const INVISIBLE_MACOS_LINE_HEIGHT_MULTI_DENSE = "1.3";

function clampYPercent(yPercent: number): number {
  return Math.max(2, Math.min(80, yPercent));
}

export function createPositioningController(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, "isAnySettingsModalOpen">;
    applySubtitleFontSize: (fontSize: number) => void;
  },
) {
  function getCurrentYPercent(): number {
    if (ctx.state.currentYPercent !== null) {
      return ctx.state.currentYPercent;
    }
    const marginBottom = parseFloat(ctx.dom.subtitleContainer.style.marginBottom) || 60;
    ctx.state.currentYPercent = clampYPercent((marginBottom / window.innerHeight) * 100);
    return ctx.state.currentYPercent;
  }

  function applyYPercent(yPercent: number): void {
    const clampedPercent = clampYPercent(yPercent);
    ctx.state.currentYPercent = clampedPercent;
    const marginBottom = (clampedPercent / 100) * window.innerHeight;

    ctx.dom.subtitleContainer.style.position = "";
    ctx.dom.subtitleContainer.style.left = "";
    ctx.dom.subtitleContainer.style.top = "";
    ctx.dom.subtitleContainer.style.right = "";
    ctx.dom.subtitleContainer.style.transform = "";

    ctx.dom.subtitleContainer.style.marginBottom = `${marginBottom}px`;
  }

  function updatePersistedSubtitlePosition(position: SubtitlePosition | null): void {
    const nextYPercent =
      position &&
      typeof position.yPercent === "number" &&
      Number.isFinite(position.yPercent)
        ? position.yPercent
        : ctx.state.persistedSubtitlePosition.yPercent;
    const nextXOffset =
      position &&
      typeof position.invisibleOffsetXPx === "number" &&
      Number.isFinite(position.invisibleOffsetXPx)
        ? position.invisibleOffsetXPx
        : 0;
    const nextYOffset =
      position &&
      typeof position.invisibleOffsetYPx === "number" &&
      Number.isFinite(position.invisibleOffsetYPx)
        ? position.invisibleOffsetYPx
        : 0;

    ctx.state.persistedSubtitlePosition = {
      yPercent: nextYPercent,
      invisibleOffsetXPx: nextXOffset,
      invisibleOffsetYPx: nextYOffset,
    };
  }

  function persistSubtitlePositionPatch(patch: Partial<SubtitlePosition>): void {
    const nextPosition: SubtitlePosition = {
      yPercent:
        typeof patch.yPercent === "number" && Number.isFinite(patch.yPercent)
          ? patch.yPercent
          : ctx.state.persistedSubtitlePosition.yPercent,
      invisibleOffsetXPx:
        typeof patch.invisibleOffsetXPx === "number" &&
        Number.isFinite(patch.invisibleOffsetXPx)
          ? patch.invisibleOffsetXPx
          : ctx.state.persistedSubtitlePosition.invisibleOffsetXPx ?? 0,
      invisibleOffsetYPx:
        typeof patch.invisibleOffsetYPx === "number" &&
        Number.isFinite(patch.invisibleOffsetYPx)
          ? patch.invisibleOffsetYPx
          : ctx.state.persistedSubtitlePosition.invisibleOffsetYPx ?? 0,
    };

    ctx.state.persistedSubtitlePosition = nextPosition;
    window.electronAPI.saveSubtitlePosition(nextPosition);
  }

  function applyStoredSubtitlePosition(
    position: SubtitlePosition | null,
    source: string,
  ): void {
    updatePersistedSubtitlePosition(position);
    if (position && position.yPercent !== undefined) {
      applyYPercent(position.yPercent);
      console.log(
        "Applied subtitle position from",
        source,
        ":",
        position.yPercent,
        "%",
      );
      return;
    }

    const defaultMarginBottom = 60;
    const defaultYPercent = (defaultMarginBottom / window.innerHeight) * 100;
    applyYPercent(defaultYPercent);
    console.log("Applied default subtitle position from", source);
  }

  function applyInvisibleSubtitleOffsetPosition(): void {
    const nextLeft =
      ctx.state.invisibleLayoutBaseLeftPx + ctx.state.invisibleSubtitleOffsetXPx;
    ctx.dom.subtitleContainer.style.left = `${nextLeft}px`;

    if (ctx.state.invisibleLayoutBaseBottomPx !== null) {
      ctx.dom.subtitleContainer.style.bottom = `${Math.max(
        0,
        ctx.state.invisibleLayoutBaseBottomPx + ctx.state.invisibleSubtitleOffsetYPx,
      )}px`;
      ctx.dom.subtitleContainer.style.top = "";
      return;
    }

    if (ctx.state.invisibleLayoutBaseTopPx !== null) {
      ctx.dom.subtitleContainer.style.top = `${Math.max(
        0,
        ctx.state.invisibleLayoutBaseTopPx - ctx.state.invisibleSubtitleOffsetYPx,
      )}px`;
      ctx.dom.subtitleContainer.style.bottom = "";
    }
  }

  function updateInvisiblePositionEditHud(): void {
    if (!ctx.state.invisiblePositionEditHud) return;
    ctx.state.invisiblePositionEditHud.textContent =
      `Position Edit  Ctrl/Cmd+Shift+P toggle  Arrow keys move  Enter/Ctrl+S save  Esc cancel  x:${Math.round(ctx.state.invisibleSubtitleOffsetXPx)} y:${Math.round(ctx.state.invisibleSubtitleOffsetYPx)}`;
  }

  function setInvisiblePositionEditMode(enabled: boolean): void {
    if (!ctx.platform.isInvisibleLayer) return;
    if (ctx.state.invisiblePositionEditMode === enabled) return;

    ctx.state.invisiblePositionEditMode = enabled;
    document.body.classList.toggle("invisible-position-edit", enabled);

    if (enabled) {
      ctx.state.invisiblePositionEditStartX = ctx.state.invisibleSubtitleOffsetXPx;
      ctx.state.invisiblePositionEditStartY = ctx.state.invisibleSubtitleOffsetYPx;
      ctx.dom.overlay.classList.add("interactive");
      if (ctx.platform.shouldToggleMouseIgnore) {
        window.electronAPI.setIgnoreMouseEvents(false);
      }
    } else if (
      !ctx.state.isOverSubtitle &&
      !options.modalStateReader.isAnySettingsModalOpen()
    ) {
      ctx.dom.overlay.classList.remove("interactive");
      if (ctx.platform.shouldToggleMouseIgnore) {
        window.electronAPI.setIgnoreMouseEvents(true, { forward: true });
      }
    }

    updateInvisiblePositionEditHud();
  }

  function applyInvisibleStoredSubtitlePosition(
    position: SubtitlePosition | null,
    source: string,
  ): void {
    updatePersistedSubtitlePosition(position);
    ctx.state.invisibleSubtitleOffsetXPx =
      ctx.state.persistedSubtitlePosition.invisibleOffsetXPx ?? 0;
    ctx.state.invisibleSubtitleOffsetYPx =
      ctx.state.persistedSubtitlePosition.invisibleOffsetYPx ?? 0;
    applyInvisibleSubtitleOffsetPosition();
    console.log(
      "[invisible-overlay] Applied subtitle offset from",
      source,
      `${ctx.state.invisibleSubtitleOffsetXPx}px`,
      `${ctx.state.invisibleSubtitleOffsetYPx}px`,
    );
    updateInvisiblePositionEditHud();
  }

  function computeOsdToCssScale(metrics: MpvSubtitleRenderMetrics): number {
    const dims = metrics.osdDimensions;
    const dpr = window.devicePixelRatio || 1;
    if (!ctx.platform.isMacOSPlatform || !dims) {
      return dpr;
    }

    const ratios = [
      dims.w / Math.max(1, window.innerWidth),
      dims.h / Math.max(1, window.innerHeight),
    ].filter((value) => Number.isFinite(value) && value > 0);

    const avgRatio =
      ratios.length > 0
        ? ratios.reduce((sum, value) => sum + value, 0) / ratios.length
        : dpr;

    return avgRatio > 1.25 ? avgRatio : 1;
  }

  function applySubtitleContainerBaseLayout(params: {
    horizontalAvailable: number;
    leftInset: number;
    marginX: number;
    hAlign: 0 | 1 | 2;
  }): void {
    ctx.dom.subtitleContainer.style.position = "absolute";
    ctx.dom.subtitleContainer.style.maxWidth = `${params.horizontalAvailable}px`;
    ctx.dom.subtitleContainer.style.width = `${params.horizontalAvailable}px`;
    ctx.dom.subtitleContainer.style.padding = "0";
    ctx.dom.subtitleContainer.style.background = "transparent";
    ctx.dom.subtitleContainer.style.marginBottom = "0";
    ctx.dom.subtitleContainer.style.pointerEvents = "none";

    ctx.dom.subtitleContainer.style.left = `${params.leftInset + params.marginX}px`;
    ctx.dom.subtitleContainer.style.right = "";
    ctx.dom.subtitleContainer.style.transform = "";
    ctx.dom.subtitleContainer.style.textAlign = "";

    if (params.hAlign === 0) {
      ctx.dom.subtitleContainer.style.textAlign = "left";
      ctx.dom.subtitleRoot.style.textAlign = "left";
    } else if (params.hAlign === 2) {
      ctx.dom.subtitleContainer.style.textAlign = "right";
      ctx.dom.subtitleRoot.style.textAlign = "right";
    } else {
      ctx.dom.subtitleContainer.style.textAlign = "center";
      ctx.dom.subtitleRoot.style.textAlign = "center";
    }

    ctx.dom.subtitleRoot.style.display = "inline-block";
    ctx.dom.subtitleRoot.style.maxWidth = "100%";
    ctx.dom.subtitleRoot.style.pointerEvents = "auto";
  }

  function applySubtitleVerticalPosition(params: {
    metrics: MpvSubtitleRenderMetrics;
    renderAreaHeight: number;
    topInset: number;
    bottomInset: number;
    marginY: number;
    effectiveFontSize: number;
    vAlign: 0 | 1 | 2;
  }): void {
    const lineCount = Math.max(1, ctx.state.currentInvisibleSubtitleLineCount);
    const multiline = lineCount > 1;
    const baselineCompensationFactor =
      lineCount >= 3 ? 0.46 : multiline ? 0.58 : 0.7;
    const baselineCompensationPx = Math.max(
      0,
      params.effectiveFontSize * baselineCompensationFactor,
    );

    if (params.vAlign === 2) {
      ctx.dom.subtitleContainer.style.top = `${Math.max(
        0,
        params.topInset + params.marginY - baselineCompensationPx,
      )}px`;
      ctx.dom.subtitleContainer.style.bottom = "";
      return;
    }

    if (params.vAlign === 1) {
      ctx.dom.subtitleContainer.style.top = "50%";
      ctx.dom.subtitleContainer.style.bottom = "";
      ctx.dom.subtitleContainer.style.transform = "translateY(-50%)";
      return;
    }

    const subPosMargin =
      ((100 - params.metrics.subPos) / 100) * params.renderAreaHeight;
    const effectiveMargin = Math.max(params.marginY, subPosMargin);
    const bottomPx = Math.max(
      0,
      params.bottomInset + effectiveMargin + baselineCompensationPx,
    );

    ctx.dom.subtitleContainer.style.top = "";
    ctx.dom.subtitleContainer.style.bottom = `${bottomPx}px`;
  }

  function applySubtitleTypography(params: {
    metrics: MpvSubtitleRenderMetrics;
    pxPerScaledPixel: number;
    effectiveFontSize: number;
  }): void {
    const lineCount = Math.max(1, ctx.state.currentInvisibleSubtitleLineCount);
    const multiline = lineCount > 1;

    ctx.dom.subtitleRoot.style.setProperty(
      "line-height",
      ctx.platform.isMacOSPlatform
        ? lineCount >= 3
          ? INVISIBLE_MACOS_LINE_HEIGHT_MULTI_DENSE
          : multiline
            ? INVISIBLE_MACOS_LINE_HEIGHT_MULTI
            : INVISIBLE_MACOS_LINE_HEIGHT_SINGLE
        : "normal",
      ctx.platform.isMacOSPlatform ? "important" : "",
    );

    const rawFont = params.metrics.subFont;
    const strippedFont = rawFont
      .replace(
        /\s+(Regular|Bold|Italic|Light|Medium|Semi\s*Bold|Extra\s*Bold|Extra\s*Light|Thin|Black|Heavy|Demi\s*Bold|Book|Condensed)\s*$/i,
        "",
      )
      .trim();

    ctx.dom.subtitleRoot.style.fontFamily =
      strippedFont !== rawFont
        ? `"${rawFont}", "${strippedFont}", sans-serif`
        : `"${rawFont}", sans-serif`;

    const effectiveSpacing = params.metrics.subSpacing;
    ctx.dom.subtitleRoot.style.setProperty(
      "letter-spacing",
      Math.abs(effectiveSpacing) > 0.0001
        ? `${effectiveSpacing * params.pxPerScaledPixel * (ctx.platform.isMacOSPlatform ? 0.7 : 1)}px`
        : ctx.platform.isMacOSPlatform
          ? "-0.02em"
          : "0px",
      ctx.platform.isMacOSPlatform ? "important" : "",
    );

    ctx.dom.subtitleRoot.style.fontKerning = ctx.platform.isMacOSPlatform
      ? "auto"
      : "none";
    ctx.dom.subtitleRoot.style.fontWeight = params.metrics.subBold ? "700" : "400";
    ctx.dom.subtitleRoot.style.fontStyle = params.metrics.subItalic
      ? "italic"
      : "normal";

    const scaleX = 1;
    const scaleY = 1;
    if (Math.abs(scaleX - 1) > 0.0001 || Math.abs(scaleY - 1) > 0.0001) {
      ctx.dom.subtitleRoot.style.transform = `scale(${scaleX}, ${scaleY})`;
      ctx.dom.subtitleRoot.style.transformOrigin = "50% 100%";
    } else {
      ctx.dom.subtitleRoot.style.transform = "";
      ctx.dom.subtitleRoot.style.transformOrigin = "";
    }

    const computedLineHeight = parseFloat(getComputedStyle(ctx.dom.subtitleRoot).lineHeight);
    if (
      Number.isFinite(computedLineHeight) &&
      computedLineHeight > params.effectiveFontSize
    ) {
      const halfLeading = (computedLineHeight - params.effectiveFontSize) / 2;
      const currentBottom = parseFloat(ctx.dom.subtitleContainer.style.bottom);
      const currentTop = parseFloat(ctx.dom.subtitleContainer.style.top);

      if (halfLeading > 0.5 && Number.isFinite(currentBottom)) {
        ctx.dom.subtitleContainer.style.bottom = `${Math.max(
          0,
          currentBottom - halfLeading,
        )}px`;
      }

      if (halfLeading > 0.5 && Number.isFinite(currentTop)) {
        ctx.dom.subtitleContainer.style.top = `${Math.max(0, currentTop - halfLeading)}px`;
      }
    }

    if (ctx.platform.isMacOSPlatform) {
      const currentBottom = parseFloat(ctx.dom.subtitleContainer.style.bottom);
      if (Number.isFinite(currentBottom)) {
        ctx.dom.subtitleContainer.style.bottom = `${Math.max(
          0,
          currentBottom + INVISIBLE_MACOS_VERTICAL_NUDGE_PX,
        )}px`;
      }
    }
  }

  function applyInvisibleSubtitleLayoutFromMpvMetrics(
    metrics: MpvSubtitleRenderMetrics,
    source: string,
  ): void {
    ctx.state.mpvSubtitleRenderMetrics = metrics;

    const dims = metrics.osdDimensions;
    const osdToCssScale = computeOsdToCssScale(metrics);
    const renderAreaHeight = dims ? dims.h / osdToCssScale : window.innerHeight;
    const renderAreaWidth = dims ? dims.w / osdToCssScale : window.innerWidth;
    const videoLeftInset = dims ? dims.ml / osdToCssScale : 0;
    const videoRightInset = dims ? dims.mr / osdToCssScale : 0;
    const videoTopInset = dims ? dims.mt / osdToCssScale : 0;
    const videoBottomInset = dims ? dims.mb / osdToCssScale : 0;

    const anchorToVideoArea = !metrics.subUseMargins;
    const leftInset = anchorToVideoArea ? videoLeftInset : 0;
    const rightInset = anchorToVideoArea ? videoRightInset : 0;
    const topInset = anchorToVideoArea ? videoTopInset : 0;
    const bottomInset = anchorToVideoArea ? videoBottomInset : 0;

    const videoHeight = renderAreaHeight - videoTopInset - videoBottomInset;
    const scaleRefHeight = metrics.subScaleByWindow ? renderAreaHeight : videoHeight;
    const pxPerScaledPixel = Math.max(0.1, scaleRefHeight / 720);
    const computedFontSize =
      metrics.subFontSize *
      metrics.subScale *
      (ctx.platform.isLinuxPlatform ? 1 : pxPerScaledPixel);

    const effectiveFontSize =
      computedFontSize * (ctx.platform.isMacOSPlatform ? 0.87 : 1);
    options.applySubtitleFontSize(effectiveFontSize);

    const marginY = metrics.subMarginY * pxPerScaledPixel;
    const marginX = Math.max(0, metrics.subMarginX * pxPerScaledPixel);
    const horizontalAvailable = Math.max(
      0,
      renderAreaWidth - leftInset - rightInset - Math.round(marginX * 2),
    );

    const effectiveBorderSize = metrics.subBorderSize * pxPerScaledPixel;
    document.documentElement.style.setProperty(
      "--sub-border-size",
      `${effectiveBorderSize}px`,
    );

    const alignment = 2;
    const hAlign = ((alignment - 1) % 3) as 0 | 1 | 2;
    const vAlign = Math.floor((alignment - 1) / 3) as 0 | 1 | 2;

    applySubtitleContainerBaseLayout({
      horizontalAvailable,
      leftInset,
      marginX,
      hAlign,
    });

    applySubtitleVerticalPosition({
      metrics,
      renderAreaHeight,
      topInset,
      bottomInset,
      marginY,
      effectiveFontSize,
      vAlign,
    });

    applySubtitleTypography({ metrics, pxPerScaledPixel, effectiveFontSize });

    ctx.state.invisibleLayoutBaseLeftPx =
      parseFloat(ctx.dom.subtitleContainer.style.left) || 0;

    const parsedBottom = parseFloat(ctx.dom.subtitleContainer.style.bottom);
    ctx.state.invisibleLayoutBaseBottomPx = Number.isFinite(parsedBottom)
      ? parsedBottom
      : null;

    const parsedTop = parseFloat(ctx.dom.subtitleContainer.style.top);
    ctx.state.invisibleLayoutBaseTopPx = Number.isFinite(parsedTop) ? parsedTop : null;

    applyInvisibleSubtitleOffsetPosition();
    updateInvisiblePositionEditHud();

    console.log(
      "[invisible-overlay] Applied mpv subtitle render metrics from",
      source,
    );
  }

  function saveInvisiblePositionEdit(): void {
    persistSubtitlePositionPatch({
      invisibleOffsetXPx: ctx.state.invisibleSubtitleOffsetXPx,
      invisibleOffsetYPx: ctx.state.invisibleSubtitleOffsetYPx,
    });
    setInvisiblePositionEditMode(false);
  }

  function cancelInvisiblePositionEdit(): void {
    ctx.state.invisibleSubtitleOffsetXPx = ctx.state.invisiblePositionEditStartX;
    ctx.state.invisibleSubtitleOffsetYPx = ctx.state.invisiblePositionEditStartY;
    applyInvisibleSubtitleOffsetPosition();
    setInvisiblePositionEditMode(false);
  }

  function setupInvisiblePositionEditHud(): void {
    if (!ctx.platform.isInvisibleLayer) return;
    const hud = document.createElement("div");
    hud.id = "invisiblePositionEditHud";
    hud.className = "invisible-position-edit-hud";
    ctx.dom.overlay.appendChild(hud);
    ctx.state.invisiblePositionEditHud = hud;
    updateInvisiblePositionEditHud();
  }

  return {
    applyInvisibleStoredSubtitlePosition,
    applyInvisibleSubtitleLayoutFromMpvMetrics,
    applyInvisibleSubtitleOffsetPosition,
    applyStoredSubtitlePosition,
    applyYPercent,
    cancelInvisiblePositionEdit,
    getCurrentYPercent,
    persistSubtitlePositionPatch,
    saveInvisiblePositionEdit,
    setInvisiblePositionEditMode,
    setupInvisiblePositionEditHud,
    updateInvisiblePositionEditHud,
  };
}
