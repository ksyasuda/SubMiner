import type { RendererContext } from './context';
import type { RendererState } from './state';
import { isYomitanPopupVisible } from './yomitan-popup.js';

function isBlockingOverlayModalOpen(state: RendererState): boolean {
  return Boolean(
    state.controllerSelectModalOpen ||
    state.controllerDebugModalOpen ||
    state.jimakuModalOpen ||
    state.youtubePickerModalOpen ||
    state.kikuModalOpen ||
    state.runtimeOptionsModalOpen ||
    state.subsyncModalOpen ||
    state.sessionHelpModalOpen,
  );
}

function isYomitanPopupInteractionActive(state: RendererState): boolean {
  if (state.yomitanPopupVisible) {
    return true;
  }
  if (typeof document === 'undefined') {
    return false;
  }
  return isYomitanPopupVisible(document);
}

export function syncOverlayMouseIgnoreState(ctx: RendererContext): void {
  const shouldKeepWindowInteractive =
    isYomitanPopupInteractionActive(ctx.state) || isBlockingOverlayModalOpen(ctx.state);
  const shouldStayInteractive =
    ctx.state.isOverSubtitle ||
    ctx.state.isOverSubtitleSidebar ||
    ctx.state.isOverYomitanPopup ||
    ctx.state.isOverOverlayNotification ||
    ctx.state.isOverNotificationHistory ||
    shouldKeepWindowInteractive;
  const shouldMarkOverlayInteractive = ctx.platform?.isLinuxPlatform
    ? shouldKeepWindowInteractive
    : shouldStayInteractive;

  if (shouldMarkOverlayInteractive) {
    ctx.dom.overlay.classList.add('interactive');
  } else {
    ctx.dom.overlay.classList.remove('interactive');
  }
  if (!ctx.platform?.shouldToggleMouseIgnore) {
    // On Linux the main process owns window passthrough via a cursor poll (Electron can't
    // forward mouse-move through a click-through window on X11). Report the interactive hint
    // only for popups/modals that sit off measured hit rects; subtitles/sidebar use the poll.
    if (ctx.platform?.isLinuxPlatform) {
      window.electronAPI.reportOverlayInteractive?.(shouldKeepWindowInteractive);
    }
    return;
  }

  if (shouldStayInteractive) {
    window.electronAPI.setIgnoreMouseEvents(false);
    return;
  }

  window.electronAPI.setIgnoreMouseEvents(true, { forward: true });
}
