import type { ModalStateReader, RendererContext } from '../context';
import { i18n } from '../../i18n/index.js';
import { resolveControllerConfigForGamepad } from '../controller-profile-config.js';

function formatAxes(values: number[]): string {
  if (values.length === 0) return i18n.t('controller.noAxes');
  return values
    .map((value, index) =>
      i18n.t('controller.debugAxis', { index: String(index), value: value.toFixed(3) }),
    )
    .join('\n');
}

function formatButtons(
  values: Array<{ value: number; pressed: boolean; touched?: boolean }>,
): string {
  if (values.length === 0) return i18n.t('controller.noButtons');
  return values
    .map((button, index) =>
      i18n.t('controller.debugButton', {
        index: String(index),
        value: button.value.toFixed(3),
        pressed: String(button.pressed),
        touched: String(button.touched ?? false),
      }),
    )
    .join('\n');
}

function formatButtonIndices(
  value: {
    select: number;
    buttonSouth: number;
    buttonEast: number;
    buttonNorth: number;
    buttonWest: number;
    leftShoulder: number;
    rightShoulder: number;
    leftStickPress: number;
    rightStickPress: number;
    leftTrigger: number;
    rightTrigger: number;
  } | null,
): string {
  if (!value) {
    return i18n.t('controller.noConfig');
  }
  return i18n.t('controller.debugButtonIndices', { value: JSON.stringify(value, null, 2) });
}

async function writeTextToClipboard(text: string): Promise<void> {
  if (!navigator.clipboard?.writeText) {
    throw new Error(i18n.t('controller.clipboardUnavailable'));
  }
  await navigator.clipboard.writeText(text);
}

export function createControllerDebugModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
    notifyControllerDisabled: () => void;
  },
) {
  let toastTimer: ReturnType<typeof setTimeout> | null = null;

  function setStatus(message: string, isError: boolean = false): void {
    ctx.dom.controllerDebugStatus.textContent = message;
    if (isError) {
      ctx.dom.controllerDebugStatus.classList.add('error');
    } else {
      ctx.dom.controllerDebugStatus.classList.remove('error');
    }
  }

  function clearToastTimer(): void {
    if (toastTimer === null) return;
    clearTimeout(toastTimer);
    toastTimer = null;
  }

  function hideToast(): void {
    clearToastTimer();
    ctx.dom.controllerDebugToast.classList.add('hidden');
    ctx.dom.controllerDebugToast.classList.remove('error');
  }

  function showToast(message: string, isError: boolean = false): void {
    clearToastTimer();
    ctx.dom.controllerDebugToast.textContent = message;
    ctx.dom.controllerDebugToast.classList.remove('hidden');
    if (isError) {
      ctx.dom.controllerDebugToast.classList.add('error');
    } else {
      ctx.dom.controllerDebugToast.classList.remove('error');
    }
    toastTimer = setTimeout(() => {
      hideToast();
    }, 1800);
  }

  function render(): void {
    const activeDevice = ctx.state.connectedGamepads.find(
      (device) => device.id === ctx.state.activeGamepadId,
    );
    setStatus(
      activeDevice?.id ??
        (ctx.state.connectedGamepads.length > 0
          ? i18n.t('controller.connected')
          : i18n.t('controller.noController')),
    );
    ctx.dom.controllerDebugSummary.textContent =
      ctx.state.connectedGamepads.length > 0
        ? ctx.state.connectedGamepads
            .map((device) => {
              const tags = [
                `#${device.index}`,
                device.mapping,
                device.id === ctx.state.activeGamepadId
                  ? i18n.t('controller.tag.active')
                  : null,
              ].filter(Boolean);
              return `${device.id || i18n.t('controller.gamepadN', { index: device.index })} (${tags.join(', ')})`;
            })
            .join('\n')
        : i18n.t('controller.connectHint');
    ctx.dom.controllerDebugAxes.textContent = formatAxes(ctx.state.controllerRawAxes);
    ctx.dom.controllerDebugButtons.textContent = formatButtons(ctx.state.controllerRawButtons);
    const activeConfig = ctx.state.controllerConfig
      ? resolveControllerConfigForGamepad(ctx.state.controllerConfig, ctx.state.activeGamepadId)
      : null;
    ctx.dom.controllerDebugButtonIndices.textContent = formatButtonIndices(
      activeConfig?.buttonIndices ?? null,
    );
  }

  async function copyButtonIndicesToClipboard(): Promise<void> {
    const text = ctx.dom.controllerDebugButtonIndices.textContent.trim();
    if (text.length === 0 || text === i18n.t('controller.noConfig')) {
      setStatus(i18n.t('controller.noButtonIndices'), true);
      showToast(i18n.t('controller.noButtonIndices'), true);
      return;
    }
    try {
      await writeTextToClipboard(text);
      setStatus(i18n.t('controller.copiedConfig'));
      showToast(i18n.t('controller.copiedConfig'));
    } catch {
      setStatus(i18n.t('controller.copyFailed'), true);
      showToast(i18n.t('controller.copyFailed'), true);
    }
  }

  function openControllerDebugModal(): boolean {
    if (ctx.state.controllerConfig?.enabled !== true) {
      options.notifyControllerDisabled();
      return false;
    }
    ctx.state.controllerDebugModalOpen = true;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.controllerDebugModal.classList.remove('hidden');
    ctx.dom.controllerDebugModal.setAttribute('aria-hidden', 'false');
    hideToast();
    render();
    return true;
  }

  function closeControllerDebugModal(): void {
    if (!ctx.state.controllerDebugModalOpen) return;
    ctx.state.controllerDebugModalOpen = false;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.controllerDebugModal.classList.add('hidden');
    ctx.dom.controllerDebugModal.setAttribute('aria-hidden', 'true');
    hideToast();
    window.electronAPI.notifyOverlayModalClosed('controller-debug');
    if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
    }
  }

  function handleControllerDebugKeydown(event: KeyboardEvent): boolean {
    if (event.key === 'Escape') {
      event.preventDefault();
      closeControllerDebugModal();
      return true;
    }
    return true;
  }

  function updateSnapshot(): void {
    if (!ctx.state.controllerDebugModalOpen) return;
    render();
  }

  function wireDomEvents(): void {
    ctx.dom.controllerDebugClose.addEventListener('click', () => {
      closeControllerDebugModal();
    });
    ctx.dom.controllerDebugCopy.addEventListener('click', () => {
      void copyButtonIndicesToClipboard();
    });
  }

  return {
    openControllerDebugModal,
    closeControllerDebugModal,
    handleControllerDebugKeydown,
    updateSnapshot,
    wireDomEvents,
  };
}
