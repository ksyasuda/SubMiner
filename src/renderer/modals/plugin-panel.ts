import type { PluginPanelSnapshot } from '../../types/runtime';
import type { ModalStateReader, RendererContext } from '../context';
import { createModalFocusGuard } from './modal-focus-guard';
import { renderPluginPanelMarkdown } from './plugin-panel-render';

const SCROLL_STEP_PX = 60;

/** Read-only panel whose content plugins push from the main process through `ctx.ui.panel`. */
export function createPluginPanelModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  let priorFocus: Element | null = null;
  let panelId: number | null = null;

  const focus = createModalFocusGuard({
    isOpen: () => ctx.state.pluginPanelModalOpen,
    getModalRoot: () => ctx.dom.pluginPanelModal,
    getPreferredFocusTargets: () => [ctx.dom.pluginPanelBody],
    getFallbackFocusTarget: () => ctx.dom.pluginPanelClose,
    isModalLayer: ctx.platform.isModalLayer,
  });

  function render(snapshot: PluginPanelSnapshot): void {
    ctx.dom.pluginPanelTitle.textContent = snapshot.title;
    renderPluginPanelMarkdown(ctx.dom.pluginPanelBody, snapshot.markdown);
  }

  function openPluginPanelModal(snapshot: PluginPanelSnapshot): void {
    panelId = snapshot.panelId;
    render(snapshot);
    ctx.dom.pluginPanelBody.scrollTop = 0;
    if (ctx.state.pluginPanelModalOpen) return;
    priorFocus = document.activeElement;

    ctx.state.pluginPanelModalOpen = true;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.pluginPanelModal.classList.remove('hidden');
    ctx.dom.pluginPanelModal.setAttribute('aria-hidden', 'false');
    ctx.dom.pluginPanelModal.setAttribute('tabindex', '-1');
    if (ctx.platform.shouldToggleMouseIgnore) {
      window.electronAPI.setIgnoreMouseEvents(false);
    }

    focus.attach();
    focus.requestOverlayFocus();
    window.focus();
    focus.enforceModalFocus();
  }

  function updatePluginPanel(snapshot: PluginPanelSnapshot): void {
    if (!ctx.state.pluginPanelModalOpen || snapshot.panelId !== panelId) return;
    render(snapshot);
  }

  function closePluginPanelModal(): void {
    if (!ctx.state.pluginPanelModalOpen) return;

    const closedPanelId = panelId ?? undefined;
    ctx.state.pluginPanelModalOpen = false;
    panelId = null;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.pluginPanelModal.classList.add('hidden');
    ctx.dom.pluginPanelModal.setAttribute('aria-hidden', 'true');
    window.electronAPI.notifyOverlayModalClosed('plugin-panel', closedPanelId);
    if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
    }

    focus.detach();

    if (priorFocus instanceof HTMLElement && priorFocus.isConnected) {
      priorFocus.focus({ preventScroll: true });
    } else if (ctx.dom.overlay instanceof HTMLElement) {
      ctx.dom.overlay.focus({ preventScroll: true });
    }

    if (ctx.platform.shouldToggleMouseIgnore) {
      if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
        window.electronAPI.setIgnoreMouseEvents(true, { forward: true });
      } else {
        window.electronAPI.setIgnoreMouseEvents(false);
      }
    }
  }

  /** The plugin closed its panel; ignore requests for a panel that was already replaced. */
  function closePluginPanelFromMain(closedPanelId: number): void {
    if (closedPanelId === panelId) closePluginPanelModal();
  }

  function handlePluginPanelKeydown(e: KeyboardEvent): boolean {
    if (!ctx.state.pluginPanelModalOpen) return false;

    if (e.key === 'Escape') {
      e.preventDefault();
      closePluginPanelModal();
      return true;
    }

    const body = ctx.dom.pluginPanelBody;
    const key = e.key.toLowerCase();
    if (key === 'arrowdown' || key === 'j') {
      e.preventDefault();
      body.scrollBy({ top: SCROLL_STEP_PX });
    } else if (key === 'arrowup' || key === 'k') {
      e.preventDefault();
      body.scrollBy({ top: -SCROLL_STEP_PX });
    } else if (key === 'pagedown' || key === ' ') {
      e.preventDefault();
      body.scrollBy({ top: body.clientHeight * 0.9 });
    } else if (key === 'pageup') {
      e.preventDefault();
      body.scrollBy({ top: -body.clientHeight * 0.9 });
    }
    return true;
  }

  function wireDomEvents(): void {
    ctx.dom.pluginPanelClose.addEventListener('click', () => {
      closePluginPanelModal();
    });
  }

  return {
    closePluginPanelModal,
    closePluginPanelFromMain,
    handlePluginPanelKeydown,
    openPluginPanelModal,
    updatePluginPanel,
    wireDomEvents,
  };
}
