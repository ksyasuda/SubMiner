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
  const shouldStayInteractive =
    ctx.state.isOverSubtitle ||
    ctx.state.isOverSubtitleSidebar ||
    isYomitanPopupInteractionActive(ctx.state) ||
    isBlockingOverlayModalOpen(ctx.state);

  if (shouldStayInteractive) {
    ctx.dom.overlay.classList.add('interactive');
  } else {
    ctx.dom.overlay.classList.remove('interactive');
  }
  if (!ctx.platform?.shouldToggleMouseIgnore) {
    return;
  }

  if (shouldStayInteractive) {
    window.electronAPI.setIgnoreMouseEvents(false);
    return;
  }

  window.electronAPI.setIgnoreMouseEvents(true, { forward: true });
}
