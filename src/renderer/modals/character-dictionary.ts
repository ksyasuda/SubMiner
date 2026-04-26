import type {
  CharacterDictionaryCandidate,
  CharacterDictionarySelectionSnapshot,
} from '../../types';
import type { ModalStateReader, RendererContext } from '../context';

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
  function setStatus(message: string, isError = false): void {
    ctx.state.characterDictionaryStatus = message;
    ctx.dom.characterDictionaryStatus.textContent = message;
    ctx.dom.characterDictionaryStatus.classList.toggle('error', isError);
  }

  function setSelection(snapshot: CharacterDictionarySelectionSnapshot): void {
    const previousId =
      ctx.state.characterDictionarySelection?.candidates[ctx.state.characterDictionarySelectedIndex]
        ?.id;
    ctx.state.characterDictionarySelection = snapshot;
    const nextIndex = snapshot.candidates.findIndex((candidate) => candidate.id === previousId);
    ctx.state.characterDictionarySelectedIndex = clampIndex(
      nextIndex >= 0 ? nextIndex : 0,
      snapshot.candidates.length,
    );
    render();
  }

  function renderCandidate(candidate: CharacterDictionaryCandidate, index: number): HTMLLIElement {
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
    button.textContent = 'Use';
    button.addEventListener('click', (event) => {
      event.stopPropagation();
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
      empty.textContent = 'No AniList candidates found.';
      ctx.dom.characterDictionaryCandidates.append(empty);
      return;
    }

    ctx.dom.characterDictionaryCandidates.replaceChildren(
      ...snapshot.candidates.map((candidate, index) => renderCandidate(candidate, index)),
    );
  }

  async function refreshSelection(): Promise<void> {
    const snapshot = await window.electronAPI.getCharacterDictionarySelection();
    setSelection(snapshot);
    setStatus(
      snapshot.override
        ? `Override active: ${formatCandidate(snapshot.override)}`
        : 'Select the correct AniList entry.',
    );
  }

  async function applySelectedCandidate(): Promise<void> {
    const snapshot = ctx.state.characterDictionarySelection;
    const candidate = snapshot?.candidates[ctx.state.characterDictionarySelectedIndex];
    if (!candidate) return;

    setStatus(`Saving override for ${candidate.title}...`);
    try {
      const result = await window.electronAPI.setCharacterDictionarySelection(candidate.id);
      if (!result.ok) {
        setStatus('Failed to save override', true);
        return;
      }
      await refreshSelection();
      const staleLabel =
        result.staleMediaIds.length > 0
          ? ` Removed stale: ${result.staleMediaIds.join(', ')}.`
          : '';
      setStatus(`Override saved: ${formatCandidate(result.selected)}.${staleLabel}`);
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
    setStatus('Loading AniList candidates...');
  }

  async function openCharacterDictionaryModal(): Promise<void> {
    if (!ctx.state.characterDictionaryModalOpen) {
      showShell();
    } else {
      setStatus('Refreshing AniList candidates...');
    }
    await refreshSelection();
  }

  function closeCharacterDictionaryModal(): void {
    if (!ctx.state.characterDictionaryModalOpen) return;
    ctx.state.characterDictionaryModalOpen = false;
    ctx.state.characterDictionarySelection = null;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.characterDictionaryModal.classList.add('hidden');
    ctx.dom.characterDictionaryModal.setAttribute('aria-hidden', 'true');
    ctx.dom.characterDictionaryCandidates.replaceChildren();
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
  }

  return {
    openCharacterDictionaryModal,
    closeCharacterDictionaryModal,
    handleCharacterDictionaryKeydown,
    wireDomEvents,
  };
}
