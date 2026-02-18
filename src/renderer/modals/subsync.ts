import type { SubsyncManualPayload } from '../../types';
import type { ModalStateReader, RendererContext } from '../context';

export function createSubsyncModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
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
    if (ctx.platform.isInvisibleLayer) return;

    ctx.state.subsyncSubmitting = false;
    ctx.dom.subsyncRunButton.disabled = false;
    ctx.state.subsyncSourceTracks = payload.sourceTracks;

    const hasSources = ctx.state.subsyncSourceTracks.length > 0;
    ctx.dom.subsyncEngineAlass.checked = hasSources;
    ctx.dom.subsyncEngineFfsubsync.checked = !hasSources;

    renderSubsyncSourceTracks();
    updateSubsyncSourceVisibility();

    setSubsyncStatus(
      hasSources
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

  async function runSubsyncManualFromModal(): Promise<void> {
    if (ctx.state.subsyncSubmitting) return;

    const engine = ctx.dom.subsyncEngineAlass.checked ? 'alass' : 'ffsubsync';
    const sourceTrackId =
      engine === 'alass' && ctx.dom.subsyncSourceSelect.value
        ? Number.parseInt(ctx.dom.subsyncSourceSelect.value, 10)
        : null;

    if (engine === 'alass' && !Number.isFinite(sourceTrackId)) {
      setSubsyncStatus('Select a source subtitle track for alass.', true);
      return;
    }

    ctx.state.subsyncSubmitting = true;
    ctx.dom.subsyncRunButton.disabled = true;

    closeSubsyncModal();
    try {
      await window.electronAPI.runSubsyncManual({
        engine,
        sourceTrackId,
      });
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
