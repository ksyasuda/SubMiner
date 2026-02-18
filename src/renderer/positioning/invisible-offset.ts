import type { SubtitlePosition } from "../../types";
import type { ModalStateReader, RendererContext } from "../context";

export type InvisibleOffsetController = {
  applyInvisibleStoredSubtitlePosition: (
    position: SubtitlePosition | null,
    source: string,
  ) => void;
  applyInvisibleSubtitleOffsetPosition: () => void;
  updateInvisiblePositionEditHud: () => void;
  setInvisiblePositionEditMode: (enabled: boolean) => void;
  saveInvisiblePositionEdit: () => void;
  cancelInvisiblePositionEdit: () => void;
  setupInvisiblePositionEditHud: () => void;
};

function formatEditHudText(offsetX: number, offsetY: number): string {
  return `Position Edit  Ctrl/Cmd+Shift+P toggle  Arrow keys move  Enter/Ctrl+S save  Esc cancel  x:${Math.round(offsetX)} y:${Math.round(offsetY)}`;
}

function createEditPositionText(ctx: RendererContext): string {
  return formatEditHudText(
    ctx.state.invisibleSubtitleOffsetXPx,
    ctx.state.invisibleSubtitleOffsetYPx,
  );
}

function applyOffsetByBasePosition(ctx: RendererContext): void {
  const nextLeft =
    ctx.state.invisibleLayoutBaseLeftPx + ctx.state.invisibleSubtitleOffsetXPx;
  ctx.dom.subtitleContainer.style.left = `${nextLeft}px`;

  if (ctx.state.invisibleLayoutBaseBottomPx !== null) {
    ctx.dom.subtitleContainer.style.bottom = `${Math.max(
      0,
      ctx.state.invisibleLayoutBaseBottomPx +
        ctx.state.invisibleSubtitleOffsetYPx,
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

export function createInvisibleOffsetController(
  ctx: RendererContext,
  modalStateReader: Pick<ModalStateReader, "isAnySettingsModalOpen">,
): InvisibleOffsetController {
  function setInvisiblePositionEditMode(enabled: boolean): void {
    if (!ctx.platform.isInvisibleLayer) return;
    if (ctx.state.invisiblePositionEditMode === enabled) return;

    ctx.state.invisiblePositionEditMode = enabled;
    document.body.classList.toggle("invisible-position-edit", enabled);

    if (enabled) {
      ctx.state.invisiblePositionEditStartX =
        ctx.state.invisibleSubtitleOffsetXPx;
      ctx.state.invisiblePositionEditStartY =
        ctx.state.invisibleSubtitleOffsetYPx;
      ctx.dom.overlay.classList.add("interactive");
      if (ctx.platform.shouldToggleMouseIgnore) {
        window.electronAPI.setIgnoreMouseEvents(false);
      }
    } else {
      if (
        !ctx.state.isOverSubtitle &&
        !modalStateReader.isAnySettingsModalOpen()
      ) {
        ctx.dom.overlay.classList.remove("interactive");
        if (ctx.platform.shouldToggleMouseIgnore) {
          window.electronAPI.setIgnoreMouseEvents(true, { forward: true });
        }
      }
    }

    updateInvisiblePositionEditHud();
  }

  function updateInvisiblePositionEditHud(): void {
    if (!ctx.state.invisiblePositionEditHud) return;
    ctx.state.invisiblePositionEditHud.textContent =
      createEditPositionText(ctx);
  }

  function applyInvisibleSubtitleOffsetPosition(): void {
    applyOffsetByBasePosition(ctx);
  }

  function applyInvisibleStoredSubtitlePosition(
    position: SubtitlePosition | null,
    source: string,
  ): void {
    if (
      position &&
      typeof position.yPercent === "number" &&
      Number.isFinite(position.yPercent)
    ) {
      ctx.state.persistedSubtitlePosition = {
        ...ctx.state.persistedSubtitlePosition,
        yPercent: position.yPercent,
      };
    }

    if (position) {
      const nextX =
        typeof position.invisibleOffsetXPx === "number" &&
        Number.isFinite(position.invisibleOffsetXPx)
          ? position.invisibleOffsetXPx
          : 0;
      const nextY =
        typeof position.invisibleOffsetYPx === "number" &&
        Number.isFinite(position.invisibleOffsetYPx)
          ? position.invisibleOffsetYPx
          : 0;
      ctx.state.invisibleSubtitleOffsetXPx = nextX;
      ctx.state.invisibleSubtitleOffsetYPx = nextY;
    } else {
      ctx.state.invisibleSubtitleOffsetXPx = 0;
      ctx.state.invisibleSubtitleOffsetYPx = 0;
    }

    applyOffsetByBasePosition(ctx);
    console.log(
      "[invisible-overlay] Applied subtitle offset from",
      source,
      `${ctx.state.invisibleSubtitleOffsetXPx}px`,
      `${ctx.state.invisibleSubtitleOffsetYPx}px`,
    );
    updateInvisiblePositionEditHud();
  }

  function saveInvisiblePositionEdit(): void {
    const nextPosition = {
      yPercent: ctx.state.persistedSubtitlePosition.yPercent,
      invisibleOffsetXPx: ctx.state.invisibleSubtitleOffsetXPx,
      invisibleOffsetYPx: ctx.state.invisibleSubtitleOffsetYPx,
    };
    window.electronAPI.saveSubtitlePosition(nextPosition);
    setInvisiblePositionEditMode(false);
  }

  function cancelInvisiblePositionEdit(): void {
    ctx.state.invisibleSubtitleOffsetXPx =
      ctx.state.invisiblePositionEditStartX;
    ctx.state.invisibleSubtitleOffsetYPx =
      ctx.state.invisiblePositionEditStartY;
    applyOffsetByBasePosition(ctx);
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
    applyInvisibleSubtitleOffsetPosition,
    updateInvisiblePositionEditHud,
    setInvisiblePositionEditMode,
    saveInvisiblePositionEdit,
    cancelInvisiblePositionEdit,
    setupInvisiblePositionEditHud,
  };
}
