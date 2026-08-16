import type { ModalStateReader, RendererContext } from '../context';
import {
  ANIME_BROWSER_CLOSE_MESSAGE,
  isAnimeBrowserKeydownMessage,
} from '../../shared/anime-browser-embed';

export function createAnimeBrowserModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    dismissOtherModals: () => void;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  function ensureFrameLoaded(): void {
    if (ctx.dom.animeBrowserFrame.src) return;
    const source = ctx.dom.animeBrowserFrame.dataset.src;
    if (!source) throw new Error('Anime Browser modal is missing its embedded source.');
    ctx.dom.animeBrowserFrame.src = source;
  }

  function openAnimeBrowserModal(): boolean {
    if (!ctx.state.animeBrowserModalOpen) {
      options.dismissOtherModals();
      if (options.modalStateReader.isAnyModalOpen()) return false;
    }

    ctx.state.animeBrowserModalOpen = true;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.animeBrowserModal.classList.remove('hidden');
    ctx.dom.animeBrowserModal.setAttribute('aria-hidden', 'false');
    ensureFrameLoaded();
    window.electronAPI.notifyOverlayModalOpened('anime-browser');
    return true;
  }

  function closeAnimeBrowserModal(): void {
    if (!ctx.state.animeBrowserModalOpen) return;
    ctx.state.animeBrowserModalOpen = false;
    ctx.dom.animeBrowserModal.classList.add('hidden');
    ctx.dom.animeBrowserModal.setAttribute('aria-hidden', 'true');
    window.electronAPI.notifyOverlayModalClosed('anime-browser');
    options.syncSettingsModalSubtitleSuppression();
    if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
    }
  }

  function handleAnimeBrowserKeydown(event: KeyboardEvent): boolean {
    if (!ctx.state.animeBrowserModalOpen || event.key !== 'Escape') return false;
    event.preventDefault();
    closeAnimeBrowserModal();
    return true;
  }

  function handleFrameMessage(event: MessageEvent): void {
    // Chromium serializes file URL message origins as "null" even though
    // Location.origin reports "file://" for these bundled pages.
    const expectedOrigin = window.location.protocol === 'file:' ? 'null' : window.location.origin;
    if (
      event.origin !== expectedOrigin ||
      event.source !== ctx.dom.animeBrowserFrame.contentWindow
    ) {
      return;
    }
    if (event.data === ANIME_BROWSER_CLOSE_MESSAGE) {
      closeAnimeBrowserModal();
      return;
    }
    if (!isAnimeBrowserKeydownMessage(event.data) || event.data.repeat) return;
    const binding = ctx.state.sessionBindingMap.get(event.data.bindingKey);
    if (binding?.actionType === 'session-action' && binding.actionId === 'openAnimeBrowser') {
      closeAnimeBrowserModal();
    }
  }

  function wireDomEvents(): void {
    ctx.dom.animeBrowserClose.addEventListener('click', closeAnimeBrowserModal);
    window.addEventListener('message', handleFrameMessage);
  }

  function disposeDomEvents(): void {
    ctx.dom.animeBrowserClose.removeEventListener('click', closeAnimeBrowserModal);
    window.removeEventListener('message', handleFrameMessage);
  }

  return {
    openAnimeBrowserModal,
    closeAnimeBrowserModal,
    handleAnimeBrowserKeydown,
    wireDomEvents,
    disposeDomEvents,
  };
}
