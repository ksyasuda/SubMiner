import type { RuntimeOptionApplyResult, RuntimeOptionState, RuntimeOptionValue } from '../../types';
import type { ModalStateReader, RendererContext } from '../context';

export function createRuntimeOptionsModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  function formatRuntimeOptionValue(value: RuntimeOptionValue): string {
    if (typeof value === 'boolean') {
      return value ? 'On' : 'Off';
    }
    return value;
  }

  function setRuntimeOptionsStatus(message: string, isError = false): void {
    ctx.dom.runtimeOptionsStatus.textContent = message;
    ctx.dom.runtimeOptionsStatus.classList.toggle('error', isError);
  }

  function getRuntimeOptionDisplayValue(option: RuntimeOptionState): RuntimeOptionValue {
    return ctx.state.runtimeOptionDraftValues.get(option.id) ?? option.value;
  }

  function getSelectedRuntimeOption(): RuntimeOptionState | null {
    if (ctx.state.runtimeOptions.length === 0) return null;
    if (ctx.state.runtimeOptionSelectedIndex < 0) return null;
    if (ctx.state.runtimeOptionSelectedIndex >= ctx.state.runtimeOptions.length) {
      return null;
    }
    return ctx.state.runtimeOptions[ctx.state.runtimeOptionSelectedIndex] ?? null;
  }

  function renderRuntimeOptionsList(): void {
    ctx.dom.runtimeOptionsList.innerHTML = '';
    ctx.state.runtimeOptions.forEach((option, index) => {
      const li = document.createElement('li');
      li.className = 'runtime-options-item';
      li.classList.toggle('active', index === ctx.state.runtimeOptionSelectedIndex);

      const label = document.createElement('div');
      label.className = 'runtime-options-label';
      label.textContent = option.label;

      const value = document.createElement('div');
      value.className = 'runtime-options-value';
      value.textContent = `Value: ${formatRuntimeOptionValue(getRuntimeOptionDisplayValue(option))}`;
      value.title = 'Click to cycle value, right-click to cycle backward';

      const allowed = document.createElement('div');
      allowed.className = 'runtime-options-allowed';
      allowed.textContent = `Allowed: ${option.allowedValues
        .map((entry) => formatRuntimeOptionValue(entry))
        .join(' | ')}`;

      li.appendChild(label);
      li.appendChild(value);
      li.appendChild(allowed);

      li.addEventListener('click', () => {
        ctx.state.runtimeOptionSelectedIndex = index;
        renderRuntimeOptionsList();
      });
      li.addEventListener('dblclick', () => {
        ctx.state.runtimeOptionSelectedIndex = index;
        void applySelectedRuntimeOption();
      });

      value.addEventListener('click', (event) => {
        event.stopPropagation();
        ctx.state.runtimeOptionSelectedIndex = index;
        cycleRuntimeDraftValue(1);
      });
      value.addEventListener('contextmenu', (event) => {
        event.preventDefault();
        event.stopPropagation();
        ctx.state.runtimeOptionSelectedIndex = index;
        cycleRuntimeDraftValue(-1);
      });

      ctx.dom.runtimeOptionsList.appendChild(li);
    });
  }

  function updateRuntimeOptions(optionsList: RuntimeOptionState[]): void {
    const previousId =
      ctx.state.runtimeOptions[ctx.state.runtimeOptionSelectedIndex]?.id ??
      ctx.state.runtimeOptions[0]?.id;

    ctx.state.runtimeOptions = optionsList;
    ctx.state.runtimeOptionDraftValues.clear();

    for (const option of ctx.state.runtimeOptions) {
      ctx.state.runtimeOptionDraftValues.set(option.id, option.value);
    }

    const nextIndex = ctx.state.runtimeOptions.findIndex((option) => option.id === previousId);
    ctx.state.runtimeOptionSelectedIndex = nextIndex >= 0 ? nextIndex : 0;

    renderRuntimeOptionsList();
  }

  function cycleRuntimeDraftValue(direction: 1 | -1): void {
    const option = getSelectedRuntimeOption();
    if (!option || option.allowedValues.length === 0) return;

    const currentValue = getRuntimeOptionDisplayValue(option);
    const currentIndex = option.allowedValues.findIndex((value) => value === currentValue);
    const safeIndex = currentIndex >= 0 ? currentIndex : 0;
    const nextIndex =
      direction === 1
        ? (safeIndex + 1) % option.allowedValues.length
        : (safeIndex - 1 + option.allowedValues.length) % option.allowedValues.length;

    const nextValue = option.allowedValues[nextIndex];
    if (nextValue === undefined) return;
    ctx.state.runtimeOptionDraftValues.set(option.id, nextValue);
    renderRuntimeOptionsList();
    setRuntimeOptionsStatus(`Selected ${option.label}: ${formatRuntimeOptionValue(nextValue)}`);
  }

  async function applySelectedRuntimeOption(): Promise<void> {
    const option = getSelectedRuntimeOption();
    if (!option) return;

    const nextValue = getRuntimeOptionDisplayValue(option);
    const result: RuntimeOptionApplyResult = await window.electronAPI.setRuntimeOptionValue(
      option.id,
      nextValue,
    );
    if (!result.ok) {
      setRuntimeOptionsStatus(result.error || 'Failed to apply option', true);
      return;
    }

    if (result.option) {
      ctx.state.runtimeOptionDraftValues.set(result.option.id, result.option.value);
    }

    const latest = await window.electronAPI.getRuntimeOptions();
    updateRuntimeOptions(latest);
    setRuntimeOptionsStatus(result.osdMessage || 'Option applied.');
  }

  function closeRuntimeOptionsModal(): void {
    if (!ctx.state.runtimeOptionsModalOpen) return;

    ctx.state.runtimeOptionsModalOpen = false;
    options.syncSettingsModalSubtitleSuppression();

    ctx.dom.runtimeOptionsModal.classList.add('hidden');
    ctx.dom.runtimeOptionsModal.setAttribute('aria-hidden', 'true');
    window.electronAPI.notifyOverlayModalClosed('runtime-options');

    setRuntimeOptionsStatus('');

    if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
    }
  }

  async function openRuntimeOptionsModal(): Promise<void> {
    const optionsList = await window.electronAPI.getRuntimeOptions();
    updateRuntimeOptions(optionsList);

    ctx.state.runtimeOptionsModalOpen = true;
    options.syncSettingsModalSubtitleSuppression();

    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.runtimeOptionsModal.classList.remove('hidden');
    ctx.dom.runtimeOptionsModal.setAttribute('aria-hidden', 'false');

    setRuntimeOptionsStatus(
      'Use arrow keys. Click value to cycle. Enter or double-click to apply.',
    );
  }

  function handleRuntimeOptionsKeydown(e: KeyboardEvent): boolean {
    if (e.key === 'Escape') {
      e.preventDefault();
      closeRuntimeOptionsModal();
      return true;
    }

    if (
      e.key === 'ArrowDown' ||
      e.key === 'j' ||
      e.key === 'J' ||
      (e.ctrlKey && (e.key === 'n' || e.key === 'N'))
    ) {
      e.preventDefault();
      if (ctx.state.runtimeOptions.length > 0) {
        ctx.state.runtimeOptionSelectedIndex = Math.min(
          ctx.state.runtimeOptions.length - 1,
          ctx.state.runtimeOptionSelectedIndex + 1,
        );
        renderRuntimeOptionsList();
      }
      return true;
    }

    if (
      e.key === 'ArrowUp' ||
      e.key === 'k' ||
      e.key === 'K' ||
      (e.ctrlKey && (e.key === 'p' || e.key === 'P'))
    ) {
      e.preventDefault();
      if (ctx.state.runtimeOptions.length > 0) {
        ctx.state.runtimeOptionSelectedIndex = Math.max(
          0,
          ctx.state.runtimeOptionSelectedIndex - 1,
        );
        renderRuntimeOptionsList();
      }
      return true;
    }

    if (e.key === 'ArrowRight' || e.key === 'l' || e.key === 'L') {
      e.preventDefault();
      cycleRuntimeDraftValue(1);
      return true;
    }

    if (e.key === 'ArrowLeft' || e.key === 'h' || e.key === 'H') {
      e.preventDefault();
      cycleRuntimeDraftValue(-1);
      return true;
    }

    if (e.key === 'Enter') {
      e.preventDefault();
      void applySelectedRuntimeOption();
      return true;
    }

    return true;
  }

  function wireDomEvents(): void {
    ctx.dom.runtimeOptionsClose.addEventListener('click', () => {
      closeRuntimeOptionsModal();
    });
  }

  return {
    closeRuntimeOptionsModal,
    handleRuntimeOptionsKeydown,
    openRuntimeOptionsModal,
    setRuntimeOptionsStatus,
    updateRuntimeOptions,
    wireDomEvents,
  };
}
