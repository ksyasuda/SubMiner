import type { ModalStateReader, RendererContext } from '../context';
import type { RuntimeOptionId, RuntimeOptionState } from '../../types/runtime-options';
import {
  buildSessionHelpSections,
  type SessionHelpSection,
  type SessionHelpTabId,
} from './session-help-sections';
import { createSessionHelpSectionNode } from './session-help-render';
import { buildVisibleSessionHelpSections, createSessionHelpTabBar } from './session-help-tabs';
import { createModalFocusGuard } from './modal-focus-guard';

export {
  buildSessionHelpSections,
  describeSessionHelpCommand,
  formatSessionHelpKeybinding,
} from './session-help-sections';

type SessionHelpBindingInfo = {
  bindingKey: 'KeyH' | 'KeyK';
  fallbackUsed: boolean;
  fallbackUnavailable: boolean;
};

/**
 * Tiers only render when known-word highlighting is also on, matching
 * getKnownWordMaturityEnabled in anki-integration/known-word-maturity.
 */
export function isKnownWordMaturityLegendEnabled(runtimeOptions: RuntimeOptionState[]): boolean {
  const isOn = (id: RuntimeOptionId): boolean =>
    runtimeOptions.some((option) => option.id === id && option.value === true);
  return (
    isOn('subtitle.annotation.knownWords.highlightEnabled') &&
    isOn('subtitle.annotation.knownWords.maturityEnabled')
  );
}

/**
 * Maturity coloring is a live runtime toggle, so the color legend reads it from
 * runtime options instead of the resolved subtitle style. A missing or failing
 * runtime-options call falls back to the flat known-word color.
 */
async function readKnownWordMaturityEnabled(): Promise<boolean> {
  try {
    return isKnownWordMaturityLegendEnabled(await window.electronAPI.getRuntimeOptions());
  } catch {
    return false;
  }
}

function formatBindingHint(info: SessionHelpBindingInfo): string {
  if (info.bindingKey === 'KeyK' && info.fallbackUsed) {
    return info.fallbackUnavailable ? 'Y-K (fallback and conflict noted)' : 'Y-K (fallback)';
  }
  return 'Y-H';
}

export function createSessionHelpModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  let priorFocus: Element | null = null;
  let openBinding: SessionHelpBindingInfo = {
    bindingKey: 'KeyH',
    fallbackUsed: false,
    fallbackUnavailable: false,
  };
  let helpFilterValue = '';
  let helpSections: SessionHelpSection[] = [];
  let activeTabId: SessionHelpTabId = 'essentials';

  function getItems(): HTMLButtonElement[] {
    return Array.from(
      ctx.dom.sessionHelpContent.querySelectorAll('.session-help-item'),
    ) as HTMLButtonElement[];
  }

  function setSelected(index: number): void {
    const items = getItems();
    if (items.length === 0) return;

    const wrappedIndex = index % items.length;
    const next = wrappedIndex < 0 ? wrappedIndex + items.length : wrappedIndex;
    ctx.state.sessionHelpSelectedIndex = next;

    items.forEach((item, idx) => {
      item.classList.toggle('active', idx === next);
      item.tabIndex = idx === next ? 0 : -1;
    });
    const activeItem = items[next];
    if (!activeItem) return;
    activeItem.focus({ preventScroll: true });
    activeItem.scrollIntoView({
      block: 'nearest',
      inline: 'nearest',
    });
  }

  const focus = createModalFocusGuard({
    isOpen: () => ctx.state.sessionHelpModalOpen,
    getModalRoot: () => ctx.dom.sessionHelpModal,
    getPreferredFocusTargets: () => getItems(),
    getFallbackFocusTarget: () => ctx.dom.sessionHelpClose,
    isModalLayer: ctx.platform.isModalLayer,
  });

  function isFilterInputFocused(): boolean {
    return document.activeElement === ctx.dom.sessionHelpFilter;
  }

  function focusFilterInput(): void {
    ctx.dom.sessionHelpFilter.focus({ preventScroll: true });
    ctx.dom.sessionHelpFilter.select();
  }

  function applyFilterAndRender(): void {
    const sections = buildVisibleSessionHelpSections(helpSections, activeTabId, helpFilterValue);
    const indexOffsets: number[] = [];
    let running = 0;
    for (const section of sections) {
      indexOffsets.push(running);
      running += section.rows.length;
    }

    ctx.dom.sessionHelpContent.innerHTML = '';
    if (!helpFilterValue.trim()) {
      ctx.dom.sessionHelpContent.appendChild(
        createSessionHelpTabBar(helpSections, activeTabId, (tabId) => {
          activeTabId = tabId;
          applyFilterAndRender();
        }),
      );
    }
    sections.forEach((section, sectionIndex) => {
      const sectionNode = createSessionHelpSectionNode(section, sectionIndex, indexOffsets);
      ctx.dom.sessionHelpContent.appendChild(sectionNode);
    });

    if (getItems().length === 0) {
      ctx.dom.sessionHelpContent.classList.add('session-help-content-no-results');
      ctx.dom.sessionHelpContent.textContent = helpFilterValue
        ? 'No matching shortcuts found.'
        : 'No active shortcuts in this tab.';
      ctx.state.sessionHelpSelectedIndex = 0;
      return;
    }

    ctx.dom.sessionHelpContent.classList.remove('session-help-content-no-results');

    if (isFilterInputFocused()) return;

    setSelected(0);
  }

  function showRenderError(message: string): void {
    helpSections = [];
    helpFilterValue = '';
    activeTabId = 'essentials';
    ctx.dom.sessionHelpFilter.value = '';
    ctx.dom.sessionHelpContent.classList.add('session-help-content-no-results');
    ctx.dom.sessionHelpContent.textContent = message;
    ctx.state.sessionHelpSelectedIndex = 0;
  }

  async function render(): Promise<boolean> {
    try {
      const [
        sessionBindings,
        styleConfig,
        markWatchedKey,
        subtitleSidebarToggleKey,
        knownWordMaturityEnabled,
      ] = await Promise.all([
        window.electronAPI.getSessionBindings(),
        window.electronAPI.getSubtitleStyle(),
        window.electronAPI.getMarkWatchedKey(),
        window.electronAPI
          .getSubtitleSidebarSnapshot()
          .then((snapshot) => snapshot.config.toggleKey)
          .catch(() => undefined),
        readKnownWordMaturityEnabled(),
      ]);

      helpSections = buildSessionHelpSections({
        sessionBindings,
        markWatchedKey,
        subtitleSidebarToggleKey,
        subtitleStyle: styleConfig ?? {},
        knownWordMaturityEnabled,
      });
      applyFilterAndRender();
      return true;
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unable to load session help data.';
      showRenderError(`Session help failed to load: ${message}`);
      return false;
    }
  }

  function openSessionHelpModal(opening: SessionHelpBindingInfo): void {
    openBinding = opening;
    priorFocus = document.activeElement;

    ctx.state.sessionHelpModalOpen = true;
    helpSections = [];
    helpFilterValue = '';
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.sessionHelpModal.classList.remove('hidden');
    ctx.dom.sessionHelpModal.setAttribute('aria-hidden', 'false');
    ctx.dom.sessionHelpModal.setAttribute('tabindex', '-1');
    ctx.dom.sessionHelpFilter.value = '';
    ctx.state.sessionHelpSelectedIndex = 0;
    ctx.dom.sessionHelpContent.innerHTML = '';
    ctx.dom.sessionHelpContent.classList.remove('session-help-content-no-results');
    if (ctx.platform.shouldToggleMouseIgnore) {
      window.electronAPI.setIgnoreMouseEvents(false);
    }

    ctx.dom.sessionHelpShortcut.textContent = `Session help opened with ${formatBindingHint(openBinding)}`;
    if (openBinding.fallbackUnavailable) {
      ctx.dom.sessionHelpWarning.textContent =
        'Both Y-H and Y-K are bound; Y-K remains the fallback for this session.';
    } else if (openBinding.fallbackUsed) {
      ctx.dom.sessionHelpWarning.textContent = 'Y-H is already bound; using Y-K as fallback.';
    } else {
      ctx.dom.sessionHelpWarning.textContent = '';
    }
    ctx.dom.sessionHelpStatus.textContent = 'Loading session help data...';

    focus.attach();
    focus.requestOverlayFocus();
    window.focus();
    focus.enforceModalFocus();

    void render().then((dataLoaded) => {
      if (!ctx.state.sessionHelpModalOpen) return;
      if (dataLoaded) {
        ctx.dom.sessionHelpStatus.textContent =
          'Use Arrow keys, J/K/H/L, mouse, click, or / then type to filter. Esc closes.';
      } else {
        ctx.dom.sessionHelpStatus.textContent =
          'Session help data is unavailable right now. Press Esc to close.';
        ctx.dom.sessionHelpWarning.textContent =
          'Unable to load latest shortcut settings from the runtime.';
      }
    });
  }

  function closeSessionHelpModal(): void {
    if (!ctx.state.sessionHelpModalOpen) return;

    ctx.state.sessionHelpModalOpen = false;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.sessionHelpModal.classList.add('hidden');
    ctx.dom.sessionHelpModal.setAttribute('aria-hidden', 'true');
    window.electronAPI.notifyOverlayModalClosed('session-help');
    if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
    }

    focus.detach();

    if (priorFocus instanceof HTMLElement && priorFocus.isConnected) {
      priorFocus.focus({ preventScroll: true });
      return;
    }

    if (ctx.dom.overlay instanceof HTMLElement) {
      // Overlay remains `tabindex="-1"` to allow programmatic focus for fallback.
      ctx.dom.overlay.focus({ preventScroll: true });
    }
    if (ctx.platform.shouldToggleMouseIgnore) {
      if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
        window.electronAPI.setIgnoreMouseEvents(true, { forward: true });
      } else {
        window.electronAPI.setIgnoreMouseEvents(false);
      }
    }
    ctx.dom.sessionHelpFilter.value = '';
    helpFilterValue = '';
    window.focus();
  }

  function handleSessionHelpKeydown(e: KeyboardEvent): boolean {
    if (!ctx.state.sessionHelpModalOpen) return false;

    if (isFilterInputFocused()) {
      if (e.key === 'Escape') {
        e.preventDefault();
        if (!helpFilterValue) {
          closeSessionHelpModal();
          return true;
        }

        helpFilterValue = '';
        ctx.dom.sessionHelpFilter.value = '';
        applyFilterAndRender();
        focus.focusFallbackTarget();
        return true;
      }
      return false;
    }

    if (e.key === 'Escape') {
      e.preventDefault();
      closeSessionHelpModal();
      return true;
    }

    const items = getItems();
    if (items.length === 0) return true;

    if (e.key === '/' && !e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey) {
      e.preventDefault();
      focusFilterInput();
      return true;
    }

    const key = e.key.toLowerCase();

    if (key === 'arrowdown' || key === 'j' || key === 'l') {
      e.preventDefault();
      setSelected(ctx.state.sessionHelpSelectedIndex + 1);
      return true;
    }

    if (key === 'arrowup' || key === 'k' || key === 'h') {
      e.preventDefault();
      setSelected(ctx.state.sessionHelpSelectedIndex - 1);
      return true;
    }

    return true;
  }

  function wireDomEvents(): void {
    ctx.dom.sessionHelpFilter.addEventListener('input', () => {
      helpFilterValue = ctx.dom.sessionHelpFilter.value;
      applyFilterAndRender();
    });

    ctx.dom.sessionHelpFilter.addEventListener('keydown', (event: KeyboardEvent) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        focus.focusFallbackTarget();
      }
    });

    ctx.dom.sessionHelpContent.addEventListener('click', (event: MouseEvent) => {
      const target = event.target;
      if (!(target instanceof Element)) return;
      const row = target.closest('.session-help-item') as HTMLElement | null;
      if (!row) return;
      const index = Number.parseInt(row.dataset.sessionHelpIndex ?? '', 10);
      if (!Number.isFinite(index)) return;
      setSelected(index);
    });

    ctx.dom.sessionHelpClose.addEventListener('click', () => {
      closeSessionHelpModal();
    });
  }

  return {
    closeSessionHelpModal,
    handleSessionHelpKeydown,
    openSessionHelpModal,
    wireDomEvents,
  };
}
