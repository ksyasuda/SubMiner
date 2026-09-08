import type { SubtitleGenerationProgress } from '../../shared/subtitle-generation';
import {
  SUBTITLE_GENERATION_MODELS,
  RECOMMENDED_SUBTITLE_GENERATION_MODEL,
  formatSubtitleGenerationModelSize,
  getSubtitleGenerationModel,
  isSubtitleGenerationModelId,
} from '../../shared/subtitle-generation-model-catalog';
import type {
  SubtitleGenerationResult,
  SubtitleGenerationStatus,
} from '../../shared/subtitle-generation-ipc';
import type { ModalStateReader, RendererContext } from '../context';
import { syncOverlayMouseIgnoreState } from '../overlay-mouse-ignore';
import { createModalFocusGuard } from './modal-focus-guard';
import {
  describeGenerationModel,
  describeGenerationProgress,
  describeGenerationVad,
} from './subtitle-generation-view';
import { SUBTITLE_GENERATION_VAD_MODEL } from '../../shared/subtitle-generation-vad-model';

function element<T extends HTMLElement>(id: string, constructor: new () => T): T {
  const node = document.getElementById(id);
  if (!(node instanceof constructor)) throw new Error(`Missing subtitle generation element: ${id}`);
  return node;
}

export function createSubtitleGenerationModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  const dom = {
    modal: element('subtitleGenerationModal', HTMLDivElement),
    close: element('subtitleGenerationClose', HTMLButtonElement),
    open: element('subtitleGenerationOpen', HTMLButtonElement),
    media: element('subtitleGenerationMedia', HTMLDivElement),
    model: element('subtitleGenerationModel', HTMLDivElement),
    modelPicker: element('subtitleGenerationModelPicker', HTMLDivElement),
    modelSelect: element('subtitleGenerationModelSelect', HTMLSelectElement),
    modelDescription: element('subtitleGenerationModelDescription', HTMLParagraphElement),
    download: element('subtitleGenerationDownload', HTMLButtonElement),
    vadEnabled: element('subtitleGenerationVadEnabled', HTMLInputElement),
    vadModel: element('subtitleGenerationVadModel', HTMLDivElement),
    vadDownload: element('subtitleGenerationVadDownload', HTMLButtonElement),
    activity: element('subtitleGenerationActivity', HTMLDivElement),
    stage: element('subtitleGenerationStage', HTMLSpanElement),
    percent: element('subtitleGenerationPercent', HTMLSpanElement),
    progress: element('subtitleGenerationProgress', HTMLProgressElement),
    status: element('subtitleGenerationStatus', HTMLDivElement),
    refresh: element('subtitleGenerationRefresh', HTMLButtonElement),
    cancel: element('subtitleGenerationCancel', HTMLButtonElement),
    start: element('subtitleGenerationStart', HTMLButtonElement),
  };
  let snapshot: SubtitleGenerationStatus | null = null;
  let progress: SubtitleGenerationProgress | null = null;
  let result: SubtitleGenerationResult | null = null;
  let pending = false;
  let cancelling = false;
  let checking = false;
  let error: string | null = null;
  let priorFocus: Element | null = null;
  let poll: ReturnType<typeof setTimeout> | null = null;
  let unsubscribe: (() => void) | null = null;
  for (const model of SUBTITLE_GENERATION_MODELS) {
    const option = document.createElement('option');
    option.value = model.id;
    const recommended = model.id === RECOMMENDED_SUBTITLE_GENERATION_MODEL ? ' (recommended)' : '';
    option.textContent = `${model.id}${recommended} · ${formatSubtitleGenerationModelSize(model.size)}`;
    dom.modelSelect.append(option);
  }

  const focus = createModalFocusGuard({
    isOpen: () => ctx.state.subtitleGenerationModalOpen,
    getModalRoot: () => dom.modal,
    getPreferredFocusTargets: () => [dom.modelSelect, dom.download, dom.start],
    getFallbackFocusTarget: () => dom.close,
    isModalLayer: ctx.platform.isModalLayer,
  });

  function render(): void {
    const busy = pending || Boolean(snapshot?.running);
    const model = snapshot ? describeGenerationModel(snapshot.model) : null;
    const vad = snapshot ? describeGenerationVad(snapshot.vad) : null;
    const readyMessage = !snapshot?.mediaPath
      ? 'Open local media to generate subtitles.'
      : !model?.ready
        ? 'Set up a speech model to continue.'
        : !vad?.ready
          ? 'Download the speech detection model or uncheck Focus on spoken dialogue.'
          : 'Ready when you are.';
    dom.media.textContent = snapshot?.mediaPath ?? 'Open a local media file in the player first.';
    dom.model.textContent = model?.text ?? 'Checking local models...';
    dom.modelPicker.classList.toggle('hidden', !snapshot || Boolean(snapshot.externalModelPath));
    dom.modelSelect.disabled = busy || checking || !snapshot || Boolean(snapshot.externalModelPath);
    if (snapshot) {
      dom.modelSelect.value = snapshot.managedModel;
      dom.modelDescription.textContent = getSubtitleGenerationModel(
        snapshot.managedModel,
      ).description;
    }
    dom.download.classList.toggle('hidden', !model?.download);
    dom.download.textContent = snapshot
      ? `Download ${snapshot.managedModel} model`
      : 'Download model';
    dom.download.disabled = busy || checking;
    if (!checking) dom.vadEnabled.checked = snapshot?.vad.enabled ?? false;
    dom.vadEnabled.disabled = busy || checking || !snapshot;
    dom.vadModel.textContent = vad?.text ?? 'Checking speech detection...';
    dom.vadDownload.classList.toggle('hidden', !vad?.download);
    dom.vadDownload.textContent = `Download speech detection model · ${formatSubtitleGenerationModelSize(SUBTITLE_GENERATION_VAD_MODEL.size)}`;
    dom.vadDownload.disabled = busy || checking;
    dom.start.disabled = busy || checking || !model?.ready || !vad?.ready || !snapshot?.mediaPath;
    dom.refresh.disabled = busy || checking;
    dom.cancel.classList.toggle('hidden', !busy);
    dom.cancel.disabled = cancelling;
    dom.cancel.textContent = cancelling ? 'Cancelling...' : 'Cancel';
    dom.activity.classList.toggle('hidden', !busy);
    const activity = describeGenerationProgress(progress);
    dom.stage.textContent = activity.stage;
    dom.percent.textContent = activity.label;
    if (activity.percent === null) dom.progress.removeAttribute('value');
    else dom.progress.value = activity.percent;
    dom.status.classList.toggle('error', Boolean(error || (!busy && result && !result.ok)));
    dom.status.textContent =
      error ??
      (busy
        ? cancelling
          ? 'Stopping the current operation...'
          : (progress?.message ?? 'Starting...')
        : (result?.message ?? (checking ? 'Checking local setup...' : readyMessage)));
  }

  function stopPolling(): void {
    if (poll !== null) clearTimeout(poll);
    poll = null;
  }

  async function refresh(): Promise<void> {
    if (checking) return;
    checking = true;
    error = null;
    render();
    try {
      snapshot = await window.electronAPI.getSubtitleGenerationStatus();
      if (!pending) {
        progress = snapshot.progress;
        result = snapshot.lastResult ?? result;
      }
    } catch (cause) {
      error = cause instanceof Error ? cause.message : 'Could not check subtitle generation setup.';
    } finally {
      checking = false;
      render();
      stopPolling();
      if (ctx.state.subtitleGenerationModalOpen && snapshot?.running && !pending) {
        poll = setTimeout(() => void refresh(), 1500);
      }
    }
  }

  async function run(action: 'download' | 'download-vad' | 'generate'): Promise<void> {
    if (pending || snapshot?.running || checking || !snapshot) return;
    const model = describeGenerationModel(snapshot.model);
    const vad = describeGenerationVad(snapshot.vad);
    if (
      action === 'download'
        ? !model.download
        : action === 'download-vad'
          ? !vad.download
          : !model.ready || !vad.ready || !snapshot.mediaPath
    )
      return;
    pending = true;
    result = null;
    error = null;
    progress = { stage: action === 'generate' ? 'extract' : 'download', message: 'Starting...' };
    render();
    try {
      result = await (action === 'download'
        ? window.electronAPI.downloadSubtitleGenerationModel()
        : action === 'download-vad'
          ? window.electronAPI.downloadSubtitleGenerationVadModel()
          : window.electronAPI.startSubtitleGeneration());
    } catch (cause) {
      result = {
        ok: false,
        message: cause instanceof Error ? cause.message : 'Subtitle generation failed.',
      };
    } finally {
      pending = false;
      cancelling = false;
      // Recheck model availability after downloads and recover controls after errors.
      await refresh();
    }
  }

  async function selectModel(): Promise<void> {
    const model = dom.modelSelect.value;
    if (pending || snapshot?.running || checking || !isSubtitleGenerationModelId(model)) return;
    checking = true;
    error = null;
    render();
    try {
      snapshot = await window.electronAPI.selectSubtitleGenerationModel(model);
      result = snapshot.lastResult;
      progress = snapshot.progress;
    } catch (cause) {
      error = cause instanceof Error ? cause.message : 'Could not select the model.';
    } finally {
      checking = false;
      render();
    }
  }

  async function selectVad(): Promise<void> {
    if (pending || snapshot?.running || checking || !snapshot) return;
    const enabled = dom.vadEnabled.checked;
    checking = true;
    error = null;
    render();
    try {
      snapshot = await window.electronAPI.setSubtitleGenerationVadEnabled(enabled);
      result = snapshot.lastResult;
      progress = snapshot.progress;
    } catch (cause) {
      error = cause instanceof Error ? cause.message : 'Could not change speech detection.';
    } finally {
      checking = false;
      render();
    }
  }

  async function cancel(): Promise<void> {
    if (cancelling) return;
    cancelling = true;
    render();
    try {
      await window.electronAPI.cancelSubtitleGeneration();
      await refresh();
    } catch (cause) {
      error = cause instanceof Error ? cause.message : 'Could not cancel the operation.';
    } finally {
      cancelling = false;
      render();
    }
  }

  function open(): void {
    if (ctx.state.subtitleGenerationModalOpen) return;
    priorFocus = document.activeElement;
    ctx.state.subtitleGenerationModalOpen = true;
    options.syncSettingsModalSubtitleSuppression();
    dom.modal.classList.remove('hidden');
    dom.modal.setAttribute('aria-hidden', 'false');
    syncOverlayMouseIgnoreState(ctx);
    focus.attach();
    focus.focusFallbackTarget();
    window.electronAPI.notifyOverlayModalOpened('subtitle-generation');
    void refresh();
  }

  function close(): void {
    if (!ctx.state.subtitleGenerationModalOpen) return;
    ctx.state.subtitleGenerationModalOpen = false;
    options.syncSettingsModalSubtitleSuppression();
    dom.modal.classList.add('hidden');
    dom.modal.setAttribute('aria-hidden', 'true');
    focus.detach();
    stopPolling();
    window.electronAPI.notifyOverlayModalClosed('subtitle-generation');
    if (priorFocus instanceof HTMLElement && priorFocus.isConnected)
      priorFocus.focus({ preventScroll: true });
    if (!options.modalStateReader.isAnyModalOpen()) syncOverlayMouseIgnoreState(ctx);
  }

  function handleKeydown(event: KeyboardEvent): boolean {
    if (!ctx.state.subtitleGenerationModalOpen) return false;
    if (event.key === 'Escape') {
      event.preventDefault();
      close();
    }
    return true;
  }

  function wireDomEvents(): void {
    dom.close.addEventListener('click', close);
    dom.open.addEventListener('click', () => {
      void window.electronAPI
        .requestSubtitleGenerationOpen()
        .then((opened) => {
          if (!opened)
            ctx.dom.subtitleSidebarStatus.textContent = 'Could not open subtitle generation.';
        })
        .catch((cause: unknown) => {
          ctx.dom.subtitleSidebarStatus.textContent =
            cause instanceof Error ? cause.message : 'Could not open subtitle generation.';
        });
    });
    dom.download.addEventListener('click', () => void run('download'));
    dom.vadDownload.addEventListener('click', () => void run('download-vad'));
    dom.vadEnabled.addEventListener('change', () => void selectVad());
    dom.modelSelect.addEventListener('change', () => void selectModel());
    dom.start.addEventListener('click', () => void run('generate'));
    dom.cancel.addEventListener('click', () => void cancel());
    dom.refresh.addEventListener('click', () => void refresh());
    unsubscribe = window.electronAPI.onSubtitleGenerationProgress((update) => {
      progress = update;
      render();
    });
  }

  function dispose(): void {
    stopPolling();
    focus.detach();
    unsubscribe?.();
    unsubscribe = null;
  }

  return { open, close, handleKeydown, wireDomEvents, dispose };
}
