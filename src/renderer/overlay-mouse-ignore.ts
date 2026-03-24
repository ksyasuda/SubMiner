import type { RendererContext } from './context';
import type { RendererState } from './state';

function isBlockingOverlayModalOpen(state: RendererState): boolean {
  const embeddedSidebarOpen =
    state.subtitleSidebarModalOpen && state.subtitleSidebarConfig?.layout === 'embedded';

  return Boolean(
    state.controllerSelectModalOpen ||
      state.controllerDebugModalOpen ||
      state.jimakuModalOpen ||
      state.youtubePickerModalOpen ||
      state.kikuModalOpen ||
      state.runtimeOptionsModalOpen ||
      state.subsyncModalOpen ||
      state.sessionHelpModalOpen ||
      (state.subtitleSidebarModalOpen && !embeddedSidebarOpen),
  );
}

export function syncOverlayMouseIgnoreState(ctx: RendererContext): void {
  const shouldStayInteractive =
    ctx.state.isOverSubtitle ||
    ctx.state.isOverSubtitleSidebar ||
    ctx.state.yomitanPopupVisible ||
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
