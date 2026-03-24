import type { YoutubePickerOpenPayload } from '../../types';
import type { ModalStateReader, RendererContext } from '../context';
import { YOMITAN_POPUP_COMMAND_EVENT } from '../yomitan-popup.js';

function createOption(value: string, label: string): HTMLOptionElement {
  const option = document.createElement('option');
  option.value = value;
  option.textContent = label;
  return option;
}

export function createYoutubeTrackPickerModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    restorePointerInteractionState: () => void;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  const OPEN_KEY_GUARD_MS = 200;
  let resolveSelectionInFlight = false;
  let keyboardSubmitEnabledAtMs = 0;

  function setStatus(message: string, isError = false): void {
    ctx.state.youtubePickerStatus = message;
    ctx.dom.youtubePickerStatus.textContent = message;
    ctx.dom.youtubePickerStatus.style.color = isError
      ? '#ed8796'
      : '#a5adcb';
  }

  function getTrackLabel(trackId: string): string {
    return ctx.state.youtubePickerPayload?.tracks.find((track) => track.id === trackId)?.label ?? '';
  }

  function renderTrackList(): void {
    ctx.dom.youtubePickerTracks.innerHTML = '';
    const payload = ctx.state.youtubePickerPayload;
    if (!payload || payload.tracks.length === 0) {
      const li = document.createElement('li');
      const left = document.createElement('span');
      left.textContent = 'No subtitle tracks found';
      const right = document.createElement('span');
      right.className = 'youtube-picker-track-meta';
      right.textContent = 'Continue without subtitles';
      li.append(left, right);
      ctx.dom.youtubePickerTracks.appendChild(li);
      return;
    }

    for (const track of payload.tracks) {
      const li = document.createElement('li');
      const left = document.createElement('span');
      left.textContent = track.label;
      const right = document.createElement('span');
      right.className = 'youtube-picker-track-meta';
      right.textContent = `${track.kind} · ${track.language}`;
      li.append(left, right);
      ctx.dom.youtubePickerTracks.appendChild(li);
    }
  }

  function setResolveControlsDisabled(disabled: boolean): void {
    ctx.dom.youtubePickerPrimarySelect.disabled = disabled;
    ctx.dom.youtubePickerSecondarySelect.disabled = disabled;
    ctx.dom.youtubePickerContinueButton.disabled = disabled;
    ctx.dom.youtubePickerCloseButton.disabled = disabled;
  }

  function syncSecondaryOptions(): void {
    const payload = ctx.state.youtubePickerPayload;
    const primaryTrackId = ctx.dom.youtubePickerPrimarySelect.value || null;
    ctx.dom.youtubePickerSecondarySelect.innerHTML = '';
    ctx.dom.youtubePickerSecondarySelect.appendChild(createOption('', 'None'));
    if (!payload) return;

    for (const track of payload.tracks) {
      if (track.id === primaryTrackId) continue;
      ctx.dom.youtubePickerSecondarySelect.appendChild(createOption(track.id, track.label));
    }
    if (
      primaryTrackId &&
      ctx.dom.youtubePickerSecondarySelect.value === primaryTrackId
    ) {
      ctx.dom.youtubePickerSecondarySelect.value = '';
    }
  }

  function setSelection(primaryTrackId: string | null, secondaryTrackId: string | null): void {
    ctx.state.youtubePickerPrimaryTrackId = primaryTrackId;
    ctx.state.youtubePickerSecondaryTrackId = secondaryTrackId;
    ctx.dom.youtubePickerPrimarySelect.value = primaryTrackId ?? '';
    syncSecondaryOptions();
    ctx.dom.youtubePickerSecondarySelect.value = secondaryTrackId ?? '';
  }

  function applyPayload(payload: YoutubePickerOpenPayload): void {
    ctx.state.youtubePickerPayload = payload;
    ctx.dom.youtubePickerTitle.textContent = `Select YouTube subtitles for ${payload.url}`;
    ctx.dom.youtubePickerPrimarySelect.innerHTML = '';
    ctx.dom.youtubePickerSecondarySelect.innerHTML = '';

    if (payload.tracks.length === 0) {
      ctx.dom.youtubePickerPrimarySelect.appendChild(createOption('', 'No tracks available'));
      ctx.dom.youtubePickerPrimarySelect.disabled = true;
      ctx.dom.youtubePickerSecondarySelect.disabled = true;
      ctx.dom.youtubePickerContinueButton.textContent = 'Continue without subtitles';
      setSelection(null, null);
      setStatus('No subtitle tracks were found. Playback will continue without subtitles.');
      renderTrackList();
      return;
    }

    ctx.dom.youtubePickerPrimarySelect.disabled = false;
    ctx.dom.youtubePickerSecondarySelect.disabled = false;
    ctx.dom.youtubePickerContinueButton.textContent = 'Use selected subtitles';
    for (const track of payload.tracks) {
      ctx.dom.youtubePickerPrimarySelect.appendChild(createOption(track.id, track.label));
    }
    setSelection(payload.defaultPrimaryTrackId, payload.defaultSecondaryTrackId);
    renderTrackList();
    setStatus('Select the subtitle tracks to download.');
  }

  async function resolveSelection(action: 'use-selected' | 'continue-without-subtitles'): Promise<void> {
    if (resolveSelectionInFlight) {
      return;
    }
    const payload = ctx.state.youtubePickerPayload;
    if (!payload) return;
    if (action === 'use-selected' && payload.hasTracks && !ctx.dom.youtubePickerPrimarySelect.value) {
      setStatus('Primary subtitle selection is required.', true);
      return;
    }

    resolveSelectionInFlight = true;
    setResolveControlsDisabled(true);
    try {
      const response =
        action === 'use-selected'
          ? await window.electronAPI.youtubePickerResolve({
              sessionId: payload.sessionId,
              action: 'use-selected',
              primaryTrackId: ctx.dom.youtubePickerPrimarySelect.value || null,
              secondaryTrackId: ctx.dom.youtubePickerSecondarySelect.value || null,
            })
          : await window.electronAPI.youtubePickerResolve({
              sessionId: payload.sessionId,
              action: 'continue-without-subtitles',
              primaryTrackId: null,
              secondaryTrackId: null,
            });
      if (!response.ok) {
        setStatus(response.message, true);
        return;
      }
      closeYoutubePickerModal();
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error), true);
    } finally {
      resolveSelectionInFlight = false;
      const shouldKeepDisabled =
        ctx.state.youtubePickerModalOpen && !(ctx.state.youtubePickerPayload?.hasTracks ?? false);
      setResolveControlsDisabled(shouldKeepDisabled);
    }
  }

  function openYoutubePickerModal(payload: YoutubePickerOpenPayload): void {
    keyboardSubmitEnabledAtMs = Date.now() + OPEN_KEY_GUARD_MS;
    if (ctx.state.youtubePickerModalOpen) {
      options.syncSettingsModalSubtitleSuppression();
      applyPayload(payload);
      window.electronAPI.notifyOverlayModalOpened('youtube-track-picker');
      return;
    }
    ctx.state.youtubePickerModalOpen = true;
    options.syncSettingsModalSubtitleSuppression();
    applyPayload(payload);
    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.youtubePickerModal.classList.remove('hidden');
    ctx.dom.youtubePickerModal.setAttribute('aria-hidden', 'false');
    window.electronAPI.notifyOverlayModalOpened('youtube-track-picker');
  }

  function closeYoutubePickerModal(): void {
    if (!ctx.state.youtubePickerModalOpen) return;
    ctx.state.youtubePickerModalOpen = false;
    options.syncSettingsModalSubtitleSuppression();
    ctx.state.youtubePickerPayload = null;
    ctx.state.youtubePickerPrimaryTrackId = null;
    ctx.state.youtubePickerSecondaryTrackId = null;
    ctx.state.youtubePickerStatus = '';
    ctx.dom.youtubePickerModal.classList.add('hidden');
    ctx.dom.youtubePickerModal.setAttribute('aria-hidden', 'true');
    window.electronAPI.notifyOverlayModalClosed('youtube-track-picker');
    window.dispatchEvent(
      new CustomEvent(YOMITAN_POPUP_COMMAND_EVENT, {
        detail: {
          type: 'refreshOptions',
        },
      }),
    );
    if (!options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
    }
    options.restorePointerInteractionState();
    void window.electronAPI.focusMainWindow();
    if (typeof ctx.dom.overlay.focus === 'function') {
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

  function handleYoutubePickerKeydown(e: KeyboardEvent): boolean {
    if (!ctx.state.youtubePickerModalOpen) return false;

    if (e.key === 'Escape') {
      e.preventDefault();
      void resolveSelection('continue-without-subtitles');
      return true;
    }

    if (e.key === 'Enter') {
      e.preventDefault();
      if (Date.now() < keyboardSubmitEnabledAtMs) {
        return true;
      }
      void resolveSelection(
        ctx.state.youtubePickerPayload?.hasTracks ? 'use-selected' : 'continue-without-subtitles',
      );
      return true;
    }

    return false;
  }

  function wireDomEvents(): void {
    ctx.dom.youtubePickerPrimarySelect.addEventListener('change', () => {
      const primaryTrackId = ctx.dom.youtubePickerPrimarySelect.value || null;
      if (ctx.dom.youtubePickerSecondarySelect.value === primaryTrackId) {
        ctx.dom.youtubePickerSecondarySelect.value = '';
      }
      setSelection(primaryTrackId, ctx.dom.youtubePickerSecondarySelect.value || null);
    });

    ctx.dom.youtubePickerSecondarySelect.addEventListener('change', () => {
      const primaryTrackId = ctx.dom.youtubePickerPrimarySelect.value || null;
      const secondaryTrackId = ctx.dom.youtubePickerSecondarySelect.value || null;
      if (primaryTrackId && secondaryTrackId === primaryTrackId) {
        ctx.dom.youtubePickerSecondarySelect.value = '';
        setStatus('Primary and secondary subtitles must be different.', true);
        return;
      }
      setSelection(primaryTrackId, secondaryTrackId);
      setStatus('Select the subtitle tracks to download.');
    });

    ctx.dom.youtubePickerContinueButton.addEventListener('click', () => {
      void resolveSelection(
        ctx.state.youtubePickerPayload?.hasTracks ? 'use-selected' : 'continue-without-subtitles',
      );
    });

    ctx.dom.youtubePickerCloseButton.addEventListener('click', () => {
      void resolveSelection('continue-without-subtitles');
    });
  }

  return {
    closeYoutubePickerModal,
    handleYoutubePickerKeydown,
    openYoutubePickerModal,
    wireDomEvents,
  };
}
