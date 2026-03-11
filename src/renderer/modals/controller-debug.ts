import type { ModalStateReader, RendererContext } from '../context';

function formatAxes(values: number[]): string {
  if (values.length === 0) return 'No controller axes available.';
  return values.map((value, index) => `axis[${index}] = ${value.toFixed(3)}`).join('\n');
}

function formatButtons(
  values: Array<{ value: number; pressed: boolean; touched?: boolean }>,
): string {
  if (values.length === 0) return 'No controller buttons available.';
  return values
    .map(
      (button, index) =>
        `button[${index}] value=${button.value.toFixed(3)} pressed=${button.pressed} touched=${button.touched ?? false}`,
    )
    .join('\n');
}

function formatButtonIndices(
  value:
    | {
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
      }
    | null,
): string {
  if (!value) {
    return 'No controller config loaded.';
  }
  return `"buttonIndices": ${JSON.stringify(value, null, 2)}`;
}

async function writeTextToClipboard(text: string): Promise<void> {
  if (!navigator.clipboard?.writeText) {
    throw new Error('Clipboard API unavailable.');
  }
  await navigator.clipboard.writeText(text);
}

export function createControllerDebugModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
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
        (ctx.state.connectedGamepads.length > 0 ? 'Controller connected.' : 'No controller detected.'),
    );
    ctx.dom.controllerDebugSummary.textContent =
      ctx.state.connectedGamepads.length > 0
        ? ctx.state.connectedGamepads
            .map((device) => {
              const tags = [
                `#${device.index}`,
                device.mapping,
                device.id === ctx.state.activeGamepadId ? 'active' : null,
              ].filter(Boolean);
              return `${device.id || `Gamepad ${device.index}`} (${tags.join(', ')})`;
            })
            .join('\n')
        : 'Connect a controller and press any button to populate raw input values.';
    ctx.dom.controllerDebugAxes.textContent = formatAxes(ctx.state.controllerRawAxes);
    ctx.dom.controllerDebugButtons.textContent = formatButtons(ctx.state.controllerRawButtons);
    ctx.dom.controllerDebugButtonIndices.textContent = formatButtonIndices(
      ctx.state.controllerConfig?.buttonIndices ?? null,
    );
  }

  async function copyButtonIndicesToClipboard(): Promise<void> {
    const text = ctx.dom.controllerDebugButtonIndices.textContent.trim();
    if (text.length === 0 || text === 'No controller config loaded.') {
      setStatus('No buttonIndices config available to copy.', true);
      showToast('No buttonIndices config available to copy.', true);
      return;
    }
    try {
      await writeTextToClipboard(text);
      setStatus('Copied controller buttonIndices config.');
      showToast('Copied controller buttonIndices config.');
    } catch {
      setStatus('Failed to copy controller buttonIndices config.', true);
      showToast('Failed to copy controller buttonIndices config.', true);
    }
  }

  function openControllerDebugModal(): void {
    ctx.state.controllerDebugModalOpen = true;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.controllerDebugModal.classList.remove('hidden');
    ctx.dom.controllerDebugModal.setAttribute('aria-hidden', 'false');
    hideToast();
    render();
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
