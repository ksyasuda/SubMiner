import type { SubsyncManualPayload } from '../../types';
import type { ModalStateReader, RendererContext } from '../context';

export function createSubsyncModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  let ffsubsyncAvailable = true;

  function setSubsyncStatus(message: string, isError = false): void {
    ctx.dom.subsyncStatus.textContent = message;
    ctx.dom.subsyncStatus.classList.toggle('error', isError);
  }

  function updateSubsyncSourceVisibility(): void {
    const useAlass = ctx.dom.subsyncEngineAlass.checked;
    ctx.dom.subsyncSourceLabel.classList.toggle('hidden', !useAlass);
  }

  function renderSubsyncSourceTracks(): void {
    ctx.dom.subsyncSourceSelect.innerHTML = '';
    for (const track of ctx.state.subsyncSourceTracks) {
      const option = document.createElement('option');
      option.value = String(track.id);
      option.textContent = track.label;
      ctx.dom.subsyncSourceSelect.appendChild(option);
    }
    ctx.dom.subsyncSourceSelect.disabled = ctx.state.subsyncSourceTracks.length === 0;
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

  function openSubsyncModal(payload: SubsyncManualPayload): void {
    ctx.state.subsyncSubmitting = false;
    ctx.state.subsyncSourceTracks = payload.sourceTracks;
    ffsubsyncAvailable = payload.ffsubsyncAvailable;

    const hasSources = ctx.state.subsyncSourceTracks.length > 0;
    ctx.dom.subsyncEngineAlass.checked = hasSources;
    ctx.dom.subsyncEngineFfsubsync.checked = !hasSources && ffsubsyncAvailable;
    ctx.dom.subsyncEngineFfsubsync.disabled = !ffsubsyncAvailable;
    ctx.dom.subsyncRunButton.disabled = !hasSources && !ffsubsyncAvailable;

    renderSubsyncSourceTracks();
    updateSubsyncSourceVisibility();

    setSubsyncStatus(
      !ffsubsyncAvailable && hasSources
        ? 'Choose alass source, then run.'
        : !ffsubsyncAvailable
          ? 'No source subtitles available for alass.'
          : hasSources
            ? 'Choose engine and source, then run.'
            : 'No source subtitles available for alass. Use ffsubsync.',
      false,
    );

    ctx.state.subsyncModalOpen = true;
    options.syncSettingsModalSubtitleSuppression();

    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.subsyncModal.classList.remove('hidden');
    ctx.dom.subsyncModal.setAttribute('aria-hidden', 'false');
  }

  function reopenSubsyncModalWithError(
    sourceTracks: SubsyncManualPayload['sourceTracks'],
    engine: 'alass' | 'ffsubsync',
    sourceTrackId: number | null,
    message: string,
  ): void {
    openSubsyncModal({ sourceTracks, ffsubsyncAvailable });

    if (engine === 'alass' && sourceTracks.length > 0) {
      ctx.dom.subsyncEngineAlass.checked = true;
      ctx.dom.subsyncEngineFfsubsync.checked = false;
      if (Number.isFinite(sourceTrackId)) {
        ctx.dom.subsyncSourceSelect.value = String(sourceTrackId);
      }
    } else if (ffsubsyncAvailable) {
      ctx.dom.subsyncEngineAlass.checked = false;
      ctx.dom.subsyncEngineFfsubsync.checked = true;
    }

    updateSubsyncSourceVisibility();
    setSubsyncStatus(message, true);
    window.electronAPI.notifyOverlayModalOpened('subsync');
  }

  async function runSubsyncManualFromModal(): Promise<void> {
    if (ctx.state.subsyncSubmitting) return;
    if (ctx.dom.subsyncRunButton.disabled) return;

    const useAlass = ctx.dom.subsyncEngineAlass.checked;
    const useFfsubsync = ctx.dom.subsyncEngineFfsubsync.checked;
    if (!useAlass && !useFfsubsync) {
      setSubsyncStatus('No sync engine available for current media.', true);
      return;
    }

    const engine = useAlass ? 'alass' : 'ffsubsync';
    const sourceTrackId =
      engine === 'alass' && ctx.dom.subsyncSourceSelect.value
        ? Number.parseInt(ctx.dom.subsyncSourceSelect.value, 10)
        : null;

    if (engine === 'alass' && !Number.isFinite(sourceTrackId)) {
      setSubsyncStatus('Select a source subtitle track for alass.', true);
      return;
    }

    const sourceTracksSnapshot = ctx.state.subsyncSourceTracks.map((track) => ({ ...track }));
    ctx.state.subsyncSubmitting = true;
    ctx.dom.subsyncRunButton.disabled = true;
    closeSubsyncModal();

    try {
      const result = await window.electronAPI.runSubsyncManual({
        engine,
        sourceTrackId,
      });
      if (result.ok) return;
      reopenSubsyncModalWithError(sourceTracksSnapshot, engine, sourceTrackId, result.message);
    } catch (error) {
      reopenSubsyncModalWithError(
        sourceTracksSnapshot,
        engine,
        sourceTrackId,
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
      updateSubsyncSourceVisibility();
    });
    ctx.dom.subsyncEngineFfsubsync.addEventListener('change', () => {
      updateSubsyncSourceVisibility();
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
