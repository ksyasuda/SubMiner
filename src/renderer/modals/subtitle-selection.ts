import type { SubtitleSelectionState } from '../../shared/subtitle-selection';
import type { RendererContext } from '../context';
import { syncOverlayMouseIgnoreState } from '../overlay-mouse-ignore';
import { createModalFocusGuard } from './modal-focus-guard';

function element<T extends HTMLElement>(id: string, constructor: new () => T): T {
  const node = document.getElementById(id);
  if (!(node instanceof constructor)) throw new Error(`Missing subtitle selection element: ${id}`);
  return node;
}

export function createSubtitleSelectionModal(
  ctx: RendererContext,
  options: { syncSettingsModalSubtitleSuppression: () => void },
) {
  const dom = {
    modal: element('subtitleSelectionModal', HTMLDivElement),
    primary: element('subtitleSelectionPrimary', HTMLSelectElement),
    secondary: element('subtitleSelectionSecondary', HTMLSelectElement),
    status: element('subtitleSelectionStatus', HTMLDivElement),
    apply: element('subtitleSelectionApply', HTMLButtonElement),
    close: element('subtitleSelectionClose', HTMLButtonElement),
  };
  let snapshot: SubtitleSelectionState | null = null;
  let generation = 0;
  let pending = false;
  let priorFocus: Element | null = null;
  const focus = createModalFocusGuard({
    isOpen: () => ctx.state.subtitleSelectionModalOpen,
    getModalRoot: () => dom.modal,
    getPreferredFocusTargets: () => [dom.primary, dom.secondary, dom.apply],
    getFallbackFocusTarget: () => dom.close,
    isModalLayer: ctx.platform.isModalLayer,
  });

  function status(message: string, error = false): void {
    dom.status.textContent = message;
    dom.status.classList.toggle('error', error);
  }

  function updateControls(): void {
    const disabled = pending || !snapshot;
    dom.primary.disabled = disabled;
    dom.secondary.disabled = disabled;
    const duplicate = dom.primary.value !== 'no' && dom.primary.value === dom.secondary.value;
    dom.apply.disabled = disabled || duplicate;
    for (const option of dom.secondary.options)
      option.disabled = option.value !== 'no' && option.value === dom.primary.value;
  }

  function populate(select: HTMLSelectElement, selected: number | null): void {
    select.replaceChildren();
    for (const track of [{ id: null, label: 'None' }, ...(snapshot?.tracks ?? [])]) {
      const option = document.createElement('option');
      option.value = track.id === null ? 'no' : String(track.id);
      option.textContent = track.label;
      select.append(option);
    }
    select.value = selected === null ? 'no' : String(selected);
  }

  async function refresh(openGeneration: number): Promise<void> {
    try {
      const next = await window.electronAPI.getSubtitleSelection();
      if (generation !== openGeneration || !ctx.state.subtitleSelectionModalOpen) return;
      snapshot = next;
      populate(dom.primary, next.primary);
      populate(dom.secondary, next.secondary);
      status(
        next.tracks.length
          ? 'Choose subtitle tracks, then apply.'
          : 'No subtitle tracks loaded in this video.',
      );
    } catch (cause) {
      if (generation !== openGeneration || !ctx.state.subtitleSelectionModalOpen) return;
      status(cause instanceof Error ? cause.message : 'Could not read subtitle tracks.', true);
    } finally {
      if (generation === openGeneration && ctx.state.subtitleSelectionModalOpen) {
        pending = false;
        updateControls();
        focus.focusFallbackTarget();
      }
    }
  }

  function open(): void {
    if (ctx.state.subtitleSelectionModalOpen) return;
    priorFocus = document.activeElement;
    snapshot = null;
    pending = true;
    generation += 1;
    populate(dom.primary, null);
    populate(dom.secondary, null);
    status('Loading subtitle tracks...');
    updateControls();
    ctx.state.subtitleSelectionModalOpen = true;
    options.syncSettingsModalSubtitleSuppression();
    dom.modal.classList.remove('hidden');
    dom.modal.setAttribute('aria-hidden', 'false');
    syncOverlayMouseIgnoreState(ctx);
    focus.attach();
    focus.focusFallbackTarget();
    window.electronAPI.notifyOverlayModalOpened('subtitle-selection');
    void refresh(generation);
  }

  function close(): void {
    if (!ctx.state.subtitleSelectionModalOpen) return;
    generation += 1;
    ctx.state.subtitleSelectionModalOpen = false;
    options.syncSettingsModalSubtitleSuppression();
    dom.modal.classList.add('hidden');
    dom.modal.setAttribute('aria-hidden', 'true');
    focus.detach();
    window.electronAPI.notifyOverlayModalClosed('subtitle-selection');
    syncOverlayMouseIgnoreState(ctx);
    if (priorFocus instanceof HTMLElement) priorFocus.focus({ preventScroll: true });
    priorFocus = null;
  }

  async function apply(): Promise<void> {
    if (dom.apply.disabled || pending || !snapshot) return;
    const openGeneration = generation;
    pending = true;
    updateControls();
    status('Applying subtitle tracks...');
    try {
      await window.electronAPI.applySubtitleSelection({
        mediaPath: snapshot.mediaPath,
        primary: dom.primary.value === 'no' ? null : Number(dom.primary.value),
        secondary: dom.secondary.value === 'no' ? null : Number(dom.secondary.value),
      });
      if (generation === openGeneration) close();
    } catch (cause) {
      if (generation === openGeneration)
        status(cause instanceof Error ? cause.message : 'Could not select subtitle tracks.', true);
    } finally {
      if (generation === openGeneration) {
        pending = false;
        updateControls();
      }
    }
  }

  function handleKeydown(event: KeyboardEvent): boolean {
    if (event.key === 'Escape') {
      event.preventDefault();
      close();
    } else if (
      event.key === 'Enter' &&
      !(event.target instanceof HTMLSelectElement) &&
      event.target !== dom.close
    ) {
      event.preventDefault();
      void apply();
    }
    return true;
  }

  function wireDomEvents(): void {
    dom.close.addEventListener('click', close);
    dom.apply.addEventListener('click', () => void apply());
    dom.primary.addEventListener('change', () => {
      if (dom.primary.value !== 'no' && dom.primary.value === dom.secondary.value)
        dom.secondary.value = 'no';
      updateControls();
    });
    dom.secondary.addEventListener('change', updateControls);
  }

  return { open, close, handleKeydown, wireDomEvents, dispose: () => focus.detach() };
}
