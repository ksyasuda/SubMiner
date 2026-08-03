import type { SubsyncManualPayload, SubsyncManualRunRequest } from '../../types';
import type { ModalStateReader, RendererContext } from '../context';

const VIDEO_REFERENCE_VALUE = 'video';

interface SubsyncSelection {
  engine: 'alass' | 'ffsubsync';
  referenceValue: string;
  targetTrackId: number | null;
}

export function createSubsyncModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  let currentPayload: SubsyncManualPayload | null = null;

  function setSubsyncStatus(message: string, isError = false): void {
    ctx.dom.subsyncStatus.textContent = message;
    ctx.dom.subsyncStatus.classList.toggle('error', isError);
  }

  function hasAlassReference(): boolean {
    if (!currentPayload) return false;
    return currentPayload.videoReferenceAvailable || currentPayload.subtitleTracks.length > 1;
  }

  function updateSubsyncFieldVisibility(): void {
    const useAlass = ctx.dom.subsyncEngineAlass.checked;
    ctx.dom.subsyncReferenceLabel.classList.toggle('hidden', !useAlass);
    ctx.dom.subsyncTargetLabel.classList.toggle(
      'hidden',
      ctx.state.subsyncSubtitleTracks.length === 0,
    );
  }

  function appendOption(select: HTMLSelectElement, value: string, label: string): void {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = label;
    select.appendChild(option);
  }

  function getSelectedTargetTrackId(): number | null {
    const raw = Number.parseInt(ctx.dom.subsyncTargetSelect.value, 10);
    return Number.isFinite(raw) ? raw : null;
  }

  function renderTargetTracks(preferredTrackId: number | null): void {
    const select = ctx.dom.subsyncTargetSelect;
    select.innerHTML = '';
    select.value = '';
    for (const track of ctx.state.subsyncSubtitleTracks) {
      appendOption(select, String(track.id), track.label);
    }
    select.disabled = ctx.state.subsyncSubtitleTracks.length === 0;

    const preferred = ctx.state.subsyncSubtitleTracks.find(
      (track) => track.id === preferredTrackId,
    );
    const fallback = ctx.state.subsyncSubtitleTracks[0];
    const selected = preferred ?? fallback;
    if (selected) {
      select.value = String(selected.id);
    }
  }

  function renderReferenceTracks(preferredValue: string | null): void {
    const select = ctx.dom.subsyncReferenceSelect;
    const targetTrackId = getSelectedTargetTrackId();
    const values: string[] = [];

    select.innerHTML = '';
    select.value = '';
    for (const track of ctx.state.subsyncSubtitleTracks) {
      if (track.id === targetTrackId) continue;
      appendOption(select, String(track.id), track.label);
      values.push(String(track.id));
    }
    if (currentPayload?.videoReferenceAvailable) {
      appendOption(select, VIDEO_REFERENCE_VALUE, 'Video file (audio reference)');
      values.push(VIDEO_REFERENCE_VALUE);
    }
    select.disabled = values.length === 0;

    const preferred = preferredValue && values.includes(preferredValue) ? preferredValue : null;
    const defaultTrackValue =
      currentPayload?.defaultReferenceTrackId !== null &&
      currentPayload?.defaultReferenceTrackId !== undefined
        ? String(currentPayload.defaultReferenceTrackId)
        : null;
    const fallback =
      defaultTrackValue && values.includes(defaultTrackValue) ? defaultTrackValue : values[0];
    const selected = preferred ?? fallback;
    if (selected) {
      select.value = selected;
    }
  }

  function describeSubsyncState(): string {
    if (!currentPayload) return '';
    const alassReady = hasAlassReference();
    if (alassReady && currentPayload.ffsubsyncAvailable) {
      return 'Choose engine, reference and out-of-sync subtitle, then run.';
    }
    if (alassReady) {
      return 'Choose the alass reference and out-of-sync subtitle, then run.';
    }
    if (currentPayload.ffsubsyncAvailable) {
      return 'No reference available for alass. Use ffsubsync.';
    }
    return 'No sync engine available for current media.';
  }

  function closeSubsyncModal(): void {
    if (!ctx.state.subsyncModalOpen) return;

    ctx.state.subsyncModalOpen = false;
    options.syncSettingsModalSubtitleSuppression();

    ctx.dom.subsyncModal.classList.add('hidden');
    ctx.dom.subsyncModal.setAttribute('aria-hidden', 'true');
    window.electronAPI.notifyOverlayModalClosed('subsync');

    if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
    }
  }

  function openSubsyncModal(payload: SubsyncManualPayload, selection?: SubsyncSelection): void {
    ctx.state.subsyncSubmitting = false;
    ctx.state.subsyncSubtitleTracks = payload.subtitleTracks;
    currentPayload = payload;

    const alassReady = hasAlassReference();
    const useAlass = selection ? selection.engine === 'alass' && alassReady : alassReady;
    ctx.dom.subsyncEngineAlass.checked = useAlass;
    ctx.dom.subsyncEngineFfsubsync.checked = !useAlass && payload.ffsubsyncAvailable;
    ctx.dom.subsyncEngineAlass.disabled = !alassReady;
    ctx.dom.subsyncEngineFfsubsync.disabled = !payload.ffsubsyncAvailable;
    ctx.dom.subsyncRunButton.disabled = !alassReady && !payload.ffsubsyncAvailable;

    renderTargetTracks(
      selection
        ? (selection.targetTrackId ?? payload.defaultTargetTrackId)
        : payload.defaultTargetTrackId,
    );
    renderReferenceTracks(selection?.referenceValue ?? null);
    updateSubsyncFieldVisibility();
    setSubsyncStatus(describeSubsyncState(), false);

    ctx.state.subsyncModalOpen = true;
    options.syncSettingsModalSubtitleSuppression();

    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.subsyncModal.classList.remove('hidden');
    ctx.dom.subsyncModal.setAttribute('aria-hidden', 'false');
  }

  function reopenSubsyncModalWithError(
    payload: SubsyncManualPayload,
    selection: SubsyncSelection,
    message: string,
  ): void {
    openSubsyncModal(payload, selection);
    setSubsyncStatus(message, true);
    window.electronAPI.notifyOverlayModalOpened('subsync');
  }

  async function runSubsyncManualFromModal(): Promise<void> {
    if (ctx.state.subsyncSubmitting) return;
    if (ctx.dom.subsyncRunButton.disabled) return;
    if (!currentPayload) return;

    const useAlass = ctx.dom.subsyncEngineAlass.checked;
    const useFfsubsync = ctx.dom.subsyncEngineFfsubsync.checked;
    if (!useAlass && !useFfsubsync) {
      setSubsyncStatus('No sync engine available for current media.', true);
      return;
    }

    const engine = useAlass ? 'alass' : 'ffsubsync';
    const referenceValue = ctx.dom.subsyncReferenceSelect.value;
    const targetTrackId = getSelectedTargetTrackId();

    if (targetTrackId === null) {
      setSubsyncStatus('Select the out-of-sync subtitle track to retime.', true);
      return;
    }
    if (engine === 'alass' && !referenceValue) {
      setSubsyncStatus('Select a reference for alass.', true);
      return;
    }

    const useVideoReference = referenceValue === VIDEO_REFERENCE_VALUE;
    const request: SubsyncManualRunRequest = {
      engine,
      targetTrackId,
    };
    if (engine === 'alass') {
      request.referenceMode = useVideoReference ? 'video' : 'track';
      request.referenceTrackId = useVideoReference ? null : Number.parseInt(referenceValue, 10);
    }

    const payloadSnapshot: SubsyncManualPayload = {
      ...currentPayload,
      subtitleTracks: currentPayload.subtitleTracks.map((track) => ({ ...track })),
    };
    const selection: SubsyncSelection = { engine, referenceValue, targetTrackId };

    ctx.state.subsyncSubmitting = true;
    ctx.dom.subsyncRunButton.disabled = true;
    closeSubsyncModal();

    try {
      const result = await window.electronAPI.runSubsyncManual(request);
      if (result.ok) return;
      reopenSubsyncModalWithError(payloadSnapshot, selection, result.message);
    } catch (error) {
      reopenSubsyncModalWithError(
        payloadSnapshot,
        selection,
        `Subsync failed: ${(error as Error).message}`,
      );
    } finally {
      ctx.state.subsyncSubmitting = false;
      ctx.dom.subsyncRunButton.disabled = false;
    }
  }

  function handleSubsyncKeydown(e: KeyboardEvent): boolean {
    if (e.key === 'Escape') {
      e.preventDefault();
      closeSubsyncModal();
      return true;
    }

    if (e.key === 'Enter') {
      e.preventDefault();
      void runSubsyncManualFromModal();
      return true;
    }

    return true;
  }

  function wireDomEvents(): void {
    ctx.dom.subsyncCloseButton.addEventListener('click', () => {
      closeSubsyncModal();
    });
    ctx.dom.subsyncEngineAlass.addEventListener('change', () => {
      updateSubsyncFieldVisibility();
    });
    ctx.dom.subsyncEngineFfsubsync.addEventListener('change', () => {
      updateSubsyncFieldVisibility();
    });
    ctx.dom.subsyncTargetSelect.addEventListener('change', () => {
      renderReferenceTracks(ctx.dom.subsyncReferenceSelect.value || null);
    });
    ctx.dom.subsyncRunButton.addEventListener('click', () => {
      void runSubsyncManualFromModal();
    });
  }

  return {
    closeSubsyncModal,
    handleSubsyncKeydown,
    openSubsyncModal,
    wireDomEvents,
  };
}
