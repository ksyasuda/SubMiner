import type { ModalStateReader, RendererContext } from '../context';

const CLOSE_MESSAGE = 'subminer:anime-browser-close';

export function createAnimeBrowserModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  function ensureFrameLoaded(): void {
    if (ctx.dom.animeBrowserFrame.src) return;
    const source = ctx.dom.animeBrowserFrame.dataset.src;
    if (!source) throw new Error('Anime Browser modal is missing its embedded source.');
    ctx.dom.animeBrowserFrame.src = source;
  }

  function openAnimeBrowserModal(): void {
    if (ctx.state.animeBrowserModalOpen || options.modalStateReader.isAnyModalOpen()) return;

    ctx.state.animeBrowserModalOpen = true;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.animeBrowserModal.classList.remove('hidden');
    ctx.dom.animeBrowserModal.setAttribute('aria-hidden', 'false');
    ensureFrameLoaded();
    window.electronAPI.notifyOverlayModalOpened('anime-browser');
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
    if (event.source === ctx.dom.animeBrowserFrame.contentWindow && event.data === CLOSE_MESSAGE) {
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
