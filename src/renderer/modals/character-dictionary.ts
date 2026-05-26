import type {
  CharacterDictionaryCandidate,
  CharacterDictionaryManagerEntry,
  CharacterDictionaryManagerSnapshot,
  CharacterDictionarySelectionSnapshot,
} from '../../types';
import type { ModalStateReader, RendererContext } from '../context';

type CharacterDictionaryView = 'override' | 'manage';

function clampIndex(index: number, length: number): number {
  if (length <= 0) return 0;
  return Math.min(Math.max(index, 0), length - 1);
}

function formatCandidate(candidate: CharacterDictionaryCandidate | null): string {
  if (!candidate) return 'None';
  const episodes = candidate.episodes === null ? '?' : String(candidate.episodes);
  return `${candidate.id} - ${candidate.title} (${episodes} episodes)`;
}

function buildSummary(snapshot: CharacterDictionarySelectionSnapshot): string {
  const guess = snapshot.guessTitle ?? 'No active title';
  return `Series key: ${snapshot.seriesKey} · Guess: ${guess}`;
}

export function createCharacterDictionaryModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  let hasSearched = false;
  let activeView: CharacterDictionaryView = 'override';
  let managerSnapshot: CharacterDictionaryManagerSnapshot | null = null;
  let pendingManagedOverride: { mediaId: number; title: string } | null = null;

  function setStatus(message: string, isError = false): void {
    ctx.state.characterDictionaryStatus = message;
    ctx.dom.characterDictionaryStatus.textContent = message;
    ctx.dom.characterDictionaryStatus.classList.toggle('error', isError);
  }

  function setSelection(
    snapshot: CharacterDictionarySelectionSnapshot,
    seedSearchInput = false,
  ): void {
    const previousId =
      ctx.state.characterDictionarySelection?.candidates[ctx.state.characterDictionarySelectedIndex]
        ?.id;
    ctx.state.characterDictionarySelection = snapshot;
    if (seedSearchInput) {
      ctx.dom.characterDictionarySearchInput.value = snapshot.guessTitle ?? '';
    }
    const nextIndex = snapshot.candidates.findIndex((candidate) => candidate.id === previousId);
    ctx.state.characterDictionarySelectedIndex = clampIndex(
      nextIndex >= 0 ? nextIndex : 0,
      snapshot.candidates.length,
    );
    render();
  }

  function setActiveView(view: CharacterDictionaryView): void {
    activeView = view;
    ctx.dom.characterDictionarySearchPanel?.classList.toggle('hidden', view !== 'override');
    ctx.dom.characterDictionaryManagerPanel?.classList.toggle('hidden', view !== 'manage');
    ctx.dom.characterDictionaryOverrideTab?.classList.toggle('active', view === 'override');
    ctx.dom.characterDictionaryManageTab?.classList.toggle('active', view === 'manage');
    ctx.dom.characterDictionaryOverrideTab?.setAttribute(
      'aria-selected',
      view === 'override' ? 'true' : 'false',
    );
    ctx.dom.characterDictionaryManageTab?.setAttribute(
      'aria-selected',
      view === 'manage' ? 'true' : 'false',
    );
  }

  function renderCandidate(candidate: CharacterDictionaryCandidate, index: number): HTMLLIElement {
    const isOverride = candidate.id === ctx.state.characterDictionarySelection?.override?.id;
    const item = document.createElement('li');
    item.className = 'character-dictionary-candidate';
    item.classList.toggle('active', index === ctx.state.characterDictionarySelectedIndex);

    const main = document.createElement('div');
    main.className = 'runtime-options-label';
    main.textContent = candidate.title;

    const meta = document.createElement('div');
    meta.className = 'runtime-options-allowed';
    const episodeLabel = candidate.episodes === null ? '?' : String(candidate.episodes);
    meta.textContent = `AniList ${candidate.id} · ${episodeLabel} episodes`;

    const button = document.createElement('button');
    button.className = 'character-dictionary-use';
    button.type = 'button';
    button.textContent = isOverride ? 'Selected' : 'Use';
    button.disabled = isOverride;
    button.addEventListener('click', (event) => {
      event.stopPropagation();
      if (isOverride) return;
      ctx.state.characterDictionarySelectedIndex = index;
      void applySelectedCandidate();
    });

    const body = document.createElement('div');
    body.className = 'character-dictionary-candidate-body';
    body.append(main, meta);

    item.append(body, button);
    item.addEventListener('click', () => {
      ctx.state.characterDictionarySelectedIndex = index;
      render();
    });
    item.addEventListener('dblclick', () => {
      ctx.state.characterDictionarySelectedIndex = index;
      void applySelectedCandidate();
    });

    return item;
  }

  function render(): void {
    const snapshot = ctx.state.characterDictionarySelection;
    ctx.dom.characterDictionaryCandidates.replaceChildren();
    if (!snapshot) {
      ctx.dom.characterDictionarySummary.textContent = '';
      ctx.dom.characterDictionaryCurrent.textContent = '';
      return;
    }

    ctx.dom.characterDictionarySummary.textContent = buildSummary(snapshot);
    ctx.dom.characterDictionaryCurrent.textContent = `Current: ${formatCandidate(
      snapshot.current,
    )} · Override: ${formatCandidate(snapshot.override)}`;

    if (snapshot.candidates.length === 0) {
      const empty = document.createElement('li');
      empty.className = 'character-dictionary-empty';
      empty.textContent = hasSearched
        ? 'No AniList candidates found.'
        : 'Search AniList to show candidates.';
      ctx.dom.characterDictionaryCandidates.append(empty);
      return;
    }

    ctx.dom.characterDictionaryCandidates.replaceChildren(
      ...snapshot.candidates.map((candidate, index) => renderCandidate(candidate, index)),
    );
  }

  function renderManagerEntry(
    entry: CharacterDictionaryManagerEntry,
    index: number,
    entryCount: number,
  ): HTMLLIElement {
    const item = document.createElement('li');
    item.className = 'character-dictionary-candidate character-dictionary-managed-entry';

    const main = document.createElement('div');
    main.className = 'runtime-options-label';
    main.textContent = entry.title || entry.label;

    const meta = document.createElement('div');
    meta.className = 'runtime-options-allowed';
    meta.textContent = `AniList ${entry.mediaId}${entry.current ? ' · Current' : ''}`;

    const body = document.createElement('div');
    body.className = 'character-dictionary-candidate-body';
    body.append(main, meta);

    const controls = document.createElement('div');
    controls.className = 'character-dictionary-manager-actions';

    const makeButton = (
      label: string,
      disabled: boolean,
      onClick: () => void | Promise<void>,
    ): HTMLButtonElement => {
      const button = document.createElement('button');
      button.className = 'character-dictionary-use';
      button.type = 'button';
      button.textContent = label;
      button.disabled = disabled;
      button.addEventListener('click', (event) => {
        event.stopPropagation();
        if (button.disabled) return;
        void onClick();
      });
      return button;
    };

    controls.append(
      makeButton('Up', entry.current || index === 0, () => moveManagedEntry(entry.mediaId, -1)),
      makeButton('Down', entry.current || index >= entryCount - 1, () =>
        moveManagedEntry(entry.mediaId, 1),
      ),
      makeButton('Override', false, () => openManagedOverride(entry)),
      makeButton('Remove', entry.current, () => removeManagedEntry(entry.mediaId)),
    );

    item.append(body, controls);
    return item;
  }

  function renderManager(): void {
    const entries = managerSnapshot?.entries ?? [];
    ctx.dom.characterDictionaryManagedEntries?.replaceChildren();
    if (!ctx.dom.characterDictionaryManagedEntries) return;

    ctx.dom.characterDictionarySummary.textContent =
      entries.length > 0
        ? `${entries.length} loaded character dictionaries. Order controls eviction priority; current dictionary stays loaded.`
        : 'No loaded character dictionaries.';
    ctx.dom.characterDictionaryCurrent.textContent = '';

    if (entries.length === 0) {
      const empty = document.createElement('li');
      empty.className = 'character-dictionary-empty';
      empty.textContent = 'No loaded character dictionaries.';
      ctx.dom.characterDictionaryManagedEntries.append(empty);
      return;
    }

    ctx.dom.characterDictionaryManagedEntries.replaceChildren(
      ...entries.map((entry, index) => renderManagerEntry(entry, index, entries.length)),
    );
  }

  async function refreshSelection(searchTitle?: string): Promise<void> {
    const snapshot = await window.electronAPI.getCharacterDictionarySelection(searchTitle);
    hasSearched = searchTitle !== '';
    setSelection(snapshot, searchTitle === '');
    setStatus(
      searchTitle === ''
        ? 'Enter a title to search AniList.'
        : snapshot.override
          ? `Override active: ${formatCandidate(snapshot.override)}`
          : 'Select the correct AniList entry.',
    );
  }

  async function refreshManager(): Promise<void> {
    managerSnapshot = await window.electronAPI.getCharacterDictionaryManagerSnapshot();
    renderManager();
    setStatus('Loaded character dictionary entries.');
  }

  async function searchCandidates(): Promise<void> {
    const searchTitle = ctx.dom.characterDictionarySearchInput.value.trim();
    if (!searchTitle) {
      setStatus('Enter a title to search AniList.', true);
      return;
    }
    ctx.dom.characterDictionarySearchButton.disabled = true;
    setStatus(`Searching AniList for ${searchTitle}...`);
    try {
      await refreshSelection(searchTitle);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error), true);
    } finally {
      ctx.dom.characterDictionarySearchButton.disabled = false;
    }
  }

  async function applySelectedCandidate(): Promise<void> {
    const snapshot = ctx.state.characterDictionarySelection;
    const candidate = snapshot?.candidates[ctx.state.characterDictionarySelectedIndex];
    if (!candidate) return;
    if (candidate.id === snapshot?.override?.id) return;

    setStatus(`Saving override for ${candidate.title}...`);
    try {
      const result = await window.electronAPI.setCharacterDictionarySelection(
        candidate.id,
        pendingManagedOverride?.mediaId,
        pendingManagedOverride ? candidate.title : undefined,
      );
      if (!result.ok) {
        setStatus('message' in result ? result.message : 'Failed to save override', true);
        return;
      }
      if (pendingManagedOverride) {
        const replacedTitle = candidate.title;
        pendingManagedOverride = null;
        await refreshManager();
        setActiveView('manage');
        setStatus(`Managed entry replaced with ${replacedTitle}.`);
        return;
      }
      await refreshSelection(ctx.dom.characterDictionarySearchInput.value.trim());
      if ('selected' in result) {
        const staleLabel =
          result.staleMediaIds.length > 0
            ? ` Removed stale: ${result.staleMediaIds.join(', ')}.`
            : '';
        setStatus(`Override saved: ${formatCandidate(result.selected)}.${staleLabel}`);
      }
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error), true);
    }
  }

  async function moveManagedEntry(mediaId: number, direction: 1 | -1): Promise<void> {
    setStatus('Updating entry order...');
    try {
      const result = await window.electronAPI.moveCharacterDictionaryManagedEntry(
        mediaId,
        direction,
      );
      managerSnapshot = { entries: result.entries };
      renderManager();
      setStatus(result.ok ? 'Entry order updated.' : result.message, !result.ok);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error), true);
    }
  }

  async function removeManagedEntry(mediaId: number): Promise<void> {
    setStatus('Removing entry...');
    try {
      const result = await window.electronAPI.removeCharacterDictionaryManagedEntry(mediaId);
      managerSnapshot = { entries: result.entries };
      renderManager();
      setStatus(
        result.ok
          ? result.rebuildRequired
            ? 'Entry removed. Merged dictionary rebuilt.'
            : 'Entry removed.'
          : result.message,
        !result.ok,
      );
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error), true);
    }
  }

  async function openManagedOverride(entry: CharacterDictionaryManagerEntry): Promise<void> {
    pendingManagedOverride = entry.current
      ? null
      : { mediaId: entry.mediaId, title: entry.title || entry.label };
    setActiveView('override');
    const searchTitle = entry.title || entry.label;
    ctx.dom.characterDictionarySearchInput.value = searchTitle;
    setStatus(`Searching AniList for ${searchTitle}...`);
    try {
      await refreshSelection(searchTitle);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error), true);
    }
  }

  function showShell(): void {
    ctx.state.characterDictionaryModalOpen = true;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.characterDictionaryModal.classList.remove('hidden');
    ctx.dom.characterDictionaryModal.setAttribute('aria-hidden', 'false');
    window.electronAPI.notifyOverlayModalOpened('character-dictionary');
    setStatus('Loading character dictionary selector...');
  }

  async function openCharacterDictionaryModal(): Promise<void> {
    setActiveView('override');
    pendingManagedOverride = null;
    if (!ctx.state.characterDictionaryModalOpen) {
      showShell();
    } else {
      window.electronAPI.notifyOverlayModalOpened('character-dictionary');
      setStatus('Refreshing AniList candidates...');
    }
    try {
      await refreshSelection('');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error), true);
    }
  }

  async function openCharacterDictionaryManagerModal(): Promise<void> {
    setActiveView('manage');
    pendingManagedOverride = null;
    if (!ctx.state.characterDictionaryModalOpen) {
      showShell();
    } else {
      window.electronAPI.notifyOverlayModalOpened('character-dictionary');
      setStatus('Refreshing character dictionary entries...');
    }
    try {
      await refreshManager();
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error), true);
    }
  }

  function closeCharacterDictionaryModal(): void {
    if (!ctx.state.characterDictionaryModalOpen) return;
    ctx.state.characterDictionaryModalOpen = false;
    ctx.state.characterDictionarySelection = null;
    managerSnapshot = null;
    pendingManagedOverride = null;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.characterDictionaryModal.classList.add('hidden');
    ctx.dom.characterDictionaryModal.setAttribute('aria-hidden', 'true');
    ctx.dom.characterDictionaryCandidates.replaceChildren();
    ctx.dom.characterDictionaryManagedEntries?.replaceChildren();
    hasSearched = false;
    window.electronAPI.notifyOverlayModalClosed('character-dictionary');
    setStatus('');
    if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
    }
  }

  function moveSelection(delta: -1 | 1): void {
    const length = ctx.state.characterDictionarySelection?.candidates.length ?? 0;
    if (length <= 0) return;
    ctx.state.characterDictionarySelectedIndex = clampIndex(
      ctx.state.characterDictionarySelectedIndex + delta,
      length,
    );
    render();
  }

  function handleCharacterDictionaryKeydown(e: KeyboardEvent): boolean {
    if (e.key === 'Escape') {
      e.preventDefault();
      closeCharacterDictionaryModal();
      return true;
    }
    if (e.target === ctx.dom.characterDictionarySearchInput) {
      if (e.key === 'Enter') {
        e.preventDefault();
        void searchCandidates();
        return true;
      }
      return false;
    }
    if (activeView === 'manage') {
      return false;
    }
    if (e.key === 'ArrowDown' || e.key === 'j' || e.key === 'J') {
      e.preventDefault();
      moveSelection(1);
      return true;
    }
    if (e.key === 'ArrowUp' || e.key === 'k' || e.key === 'K') {
      e.preventDefault();
      moveSelection(-1);
      return true;
    }
    if (e.key === 'Enter') {
      e.preventDefault();
      void applySelectedCandidate();
      return true;
    }
    return false;
  }

  function wireDomEvents(): void {
    ctx.dom.characterDictionaryClose.addEventListener('click', closeCharacterDictionaryModal);
    ctx.dom.characterDictionaryOverrideTab?.addEventListener('click', () => {
      void openCharacterDictionaryModal();
    });
    ctx.dom.characterDictionaryManageTab?.addEventListener('click', () => {
      void openCharacterDictionaryManagerModal();
    });
    ctx.dom.characterDictionarySearchButton.addEventListener('click', () => {
      void searchCandidates();
    });
    ctx.dom.characterDictionarySearchInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        void searchCandidates();
      }
    });
  }

  return {
    openCharacterDictionaryModal,
    openCharacterDictionaryManagerModal,
    closeCharacterDictionaryModal,
    handleCharacterDictionaryKeydown,
    wireDomEvents,
  };
}
