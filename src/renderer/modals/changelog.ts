import type { ChangelogSnapshot } from '../../types/changelog';
import type { ModalStateReader, RendererContext } from '../context';
import {
  createChangelogEntryNode,
  describeChangelogSource,
  resolveEntryBadge,
  shouldEntryStartExpanded,
} from './changelog-render';
import { createModalFocusGuard } from './modal-focus-guard';

export function createChangelogModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  let priorFocus: Element | null = null;
  let loadToken = 0;

  function getSummaries(): HTMLElement[] {
    return Array.from(
      ctx.dom.changelogList.querySelectorAll('.changelog-entry-summary'),
    ) as HTMLElement[];
  }

  function applySelectionStyles(index: number): HTMLElement[] {
    const summaries = getSummaries();
    ctx.state.changelogSelectedIndex = index;
    summaries.forEach((summary, idx) => {
      summary.classList.toggle('active', idx === index);
      summary.tabIndex = idx === index ? 0 : -1;
    });
    return summaries;
  }

  function setSelected(index: number): void {
    const count = getSummaries().length;
    if (count === 0) return;

    const wrapped = index % count;
    const next = wrapped < 0 ? wrapped + count : wrapped;

    // Only the keyboard path moves focus; clicking already focused the summary.
    const active = applySelectionStyles(next)[next];
    if (!active) return;
    active.focus({ preventScroll: true });
    active.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  function getSelectedEntry(): HTMLDetailsElement | null {
    const summary = getSummaries()[ctx.state.changelogSelectedIndex];
    const entry = summary?.parentElement;
    return entry instanceof HTMLDetailsElement ? entry : null;
  }

  const focus = createModalFocusGuard({
    isOpen: () => ctx.state.changelogModalOpen,
    getModalRoot: () => ctx.dom.changelogModal,
    getPreferredFocusTargets: () => getSummaries(),
    getFallbackFocusTarget: () => ctx.dom.changelogClose,
    isModalLayer: ctx.platform.isModalLayer,
  });

  function renderSnapshot(snapshot: ChangelogSnapshot): void {
    ctx.dom.changelogList.innerHTML = '';
    ctx.dom.changelogList.classList.remove('changelog-list-empty');

    ctx.dom.changelogInstalled.textContent = `Installed v${snapshot.installedVersion}`;
    ctx.dom.changelogSource.textContent = describeChangelogSource(snapshot);
    ctx.dom.changelogWarning.textContent = snapshot.warning ?? '';

    if (snapshot.entries.length === 0) {
      ctx.dom.changelogList.classList.add('changelog-list-empty');
      ctx.dom.changelogList.textContent =
        snapshot.error ?? 'No changelog entries are available right now.';
      ctx.state.changelogSelectedIndex = 0;
      return;
    }

    snapshot.entries.forEach((entry, index) => {
      ctx.dom.changelogList.appendChild(
        createChangelogEntryNode(entry, {
          expanded: shouldEntryStartExpanded(entry, snapshot),
          badge: resolveEntryBadge(entry.version, snapshot.installedVersion),
          index,
        }),
      );
    });

    setSelected(0);
  }

  async function load(options?: { refresh?: boolean }): Promise<void> {
    const token = ++loadToken;
    ctx.dom.changelogStatus.textContent = options?.refresh
      ? 'Refreshing changelog...'
      : 'Loading changelog...';

    try {
      const snapshot = await window.electronAPI.getChangelogSnapshot(options);
      if (token !== loadToken || !ctx.state.changelogModalOpen) return;
      renderSnapshot(snapshot);
      ctx.dom.changelogStatus.textContent =
        'J/K or arrows to move, Enter to fold, R to refresh, Esc closes.';
    } catch (error) {
      if (token !== loadToken || !ctx.state.changelogModalOpen) return;
      const message = error instanceof Error ? error.message : 'Unknown error.';
      ctx.dom.changelogList.innerHTML = '';
      ctx.dom.changelogList.classList.add('changelog-list-empty');
      // A failed refresh replaces an already-rendered snapshot, so the metadata
      // line has to be cleared here too or it keeps describing stale entries.
      ctx.dom.changelogInstalled.textContent = '';
      ctx.dom.changelogSource.textContent = '';
      ctx.dom.changelogWarning.textContent = '';
      ctx.dom.changelogList.textContent = `Changelog failed to load: ${message}`;
      ctx.dom.changelogStatus.textContent = 'Press Esc to close.';
    }
  }

  function openChangelogModal(): void {
    if (ctx.state.changelogModalOpen) return;
    priorFocus = document.activeElement;

    ctx.state.changelogModalOpen = true;
    ctx.state.changelogSelectedIndex = 0;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.changelogModal.classList.remove('hidden');
    ctx.dom.changelogModal.setAttribute('aria-hidden', 'false');
    ctx.dom.changelogModal.setAttribute('tabindex', '-1');
    ctx.dom.changelogList.innerHTML = '';
    ctx.dom.changelogWarning.textContent = '';
    // Reset the metadata line too, so a failed load can't leave the previous
    // session's installed/source values on screen.
    ctx.dom.changelogInstalled.textContent = '';
    ctx.dom.changelogSource.textContent = '';
    if (ctx.platform.shouldToggleMouseIgnore) {
      window.electronAPI.setIgnoreMouseEvents(false);
    }

    focus.attach();
    focus.requestOverlayFocus();
    window.focus();
    focus.enforceModalFocus();

    void load();
  }

  function closeChangelogModal(): void {
    if (!ctx.state.changelogModalOpen) return;

    ctx.state.changelogModalOpen = false;
    loadToken += 1;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.changelogModal.classList.add('hidden');
    ctx.dom.changelogModal.setAttribute('aria-hidden', 'true');
    window.electronAPI.notifyOverlayModalClosed('changelog');
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
    window.focus();
  }

  function handleChangelogKeydown(e: KeyboardEvent): boolean {
    if (!ctx.state.changelogModalOpen) return false;

    if (e.key === 'Escape') {
      e.preventDefault();
      closeChangelogModal();
      return true;
    }

    const key = e.key.toLowerCase();

    if (key === 'r' && !e.ctrlKey && !e.metaKey && !e.altKey) {
      e.preventDefault();
      void load({ refresh: true });
      return true;
    }

    const summaries = getSummaries();
    if (summaries.length === 0) return true;

    if (key === 'arrowdown' || key === 'j') {
      e.preventDefault();
      setSelected(ctx.state.changelogSelectedIndex + 1);
      return true;
    }

    if (key === 'arrowup' || key === 'k') {
      e.preventDefault();
      setSelected(ctx.state.changelogSelectedIndex - 1);
      return true;
    }

    if (key === 'enter' || key === ' ') {
      // Only the selected release summary folds from here. The Close/Refresh
      // buttons and the nested "Internal changes" fold activate themselves, and
      // swallowing Enter/Space would make them unreachable by keyboard.
      if (e.target !== summaries[ctx.state.changelogSelectedIndex]) {
        return true;
      }
      e.preventDefault();
      const entry = getSelectedEntry();
      if (entry) entry.open = !entry.open;
      return true;
    }

    if (key === 'arrowleft' || key === 'h') {
      e.preventDefault();
      const entry = getSelectedEntry();
      if (entry) entry.open = false;
      return true;
    }

    if (key === 'arrowright' || key === 'l') {
      e.preventDefault();
      const entry = getSelectedEntry();
      if (entry) entry.open = true;
      return true;
    }

    return true;
  }

  function wireDomEvents(): void {
    ctx.dom.changelogClose.addEventListener('click', () => {
      closeChangelogModal();
    });

    ctx.dom.changelogRefresh.addEventListener('click', () => {
      void load({ refresh: true });
    });

    ctx.dom.changelogList.addEventListener('click', (event: MouseEvent) => {
      const target = event.target;
      if (!(target instanceof Element)) return;
      const summary = target.closest('.changelog-entry-summary') as HTMLElement | null;
      if (!summary) return;
      const index = Number.parseInt(summary.dataset.changelogIndex ?? '', 10);
      if (!Number.isFinite(index)) return;
      applySelectionStyles(index);
    });
  }

  return {
    closeChangelogModal,
    handleChangelogKeydown,
    openChangelogModal,
    wireDomEvents,
  };
}
