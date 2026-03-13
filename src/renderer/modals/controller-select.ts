import type { ModalStateReader, RendererContext } from '../context';
import { createControllerBindingCapture } from '../handlers/controller-binding-capture.js';
import {
  createControllerConfigForm,
  getControllerBindingDefinition,
  getDefaultControllerBinding,
} from './controller-config-form.js';

function clampSelectedIndex(ctx: RendererContext): void {
  if (ctx.state.connectedGamepads.length === 0) {
    ctx.state.controllerDeviceSelectedIndex = 0;
    return;
  }

  ctx.state.controllerDeviceSelectedIndex = Math.min(
    Math.max(ctx.state.controllerDeviceSelectedIndex, 0),
    ctx.state.connectedGamepads.length - 1,
  );
}

export function createControllerSelectModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  let selectedControllerId: string | null = null;
  let lastRenderedDevicesKey = '';
  let lastRenderedActiveGamepadId: string | null = null;
  let lastRenderedPreferredId = '';
  let learningActionId: keyof NonNullable<typeof ctx.state.controllerConfig>['bindings'] | null = null;
  let bindingCapture: ReturnType<typeof createControllerBindingCapture> | null = null;

  const controllerConfigForm = createControllerConfigForm({
    container: ctx.dom.controllerConfigList,
    getBindings: () =>
      ctx.state.controllerConfig?.bindings ?? {
        toggleLookup: { kind: 'button', buttonIndex: 0 },
        closeLookup: { kind: 'button', buttonIndex: 1 },
        toggleKeyboardOnlyMode: { kind: 'button', buttonIndex: 3 },
        mineCard: { kind: 'button', buttonIndex: 2 },
        quitMpv: { kind: 'button', buttonIndex: 6 },
        previousAudio: { kind: 'none' },
        nextAudio: { kind: 'button', buttonIndex: 5 },
        playCurrentAudio: { kind: 'button', buttonIndex: 4 },
        toggleMpvPause: { kind: 'button', buttonIndex: 9 },
        leftStickHorizontal: { kind: 'axis', axisIndex: 0, dpadFallback: 'horizontal' },
        leftStickVertical: { kind: 'axis', axisIndex: 1, dpadFallback: 'vertical' },
        rightStickHorizontal: { kind: 'axis', axisIndex: 3, dpadFallback: 'none' },
        rightStickVertical: { kind: 'axis', axisIndex: 4, dpadFallback: 'none' },
      },
    getLearningActionId: () => learningActionId,
    onLearn: (actionId, bindingType) => {
      const definition = getControllerBindingDefinition(actionId);
      if (!definition) return;
      const config = ctx.state.controllerConfig;
      bindingCapture = createControllerBindingCapture({
        triggerDeadzone: config?.triggerDeadzone ?? 0.5,
        stickDeadzone: config?.stickDeadzone ?? 0.2,
      });
      bindingCapture.arm(
        bindingType === 'axis'
          ? {
              actionId,
              bindingType: 'axis',
              dpadFallback:
                definition.defaultBinding.kind === 'axis' &&
                'dpadFallback' in definition.defaultBinding
                  ? definition.defaultBinding.dpadFallback
                  : 'none',
            }
          : {
              actionId,
              bindingType: 'discrete',
            },
        {
          axes: ctx.state.controllerRawAxes,
          buttons: ctx.state.controllerRawButtons,
        },
      );
      learningActionId = actionId;
      controllerConfigForm.render();
      setStatus(`Waiting for input for ${definition.label}.`);
    },
    onClear: (actionId) => {
      void saveBinding(actionId, { kind: 'none' });
    },
    onReset: (actionId) => {
      void saveBinding(actionId, getDefaultControllerBinding(actionId));
    },
  });

  function getDevicesKey(): string {
    return ctx.state.connectedGamepads
      .map((device) => `${device.id}|${device.index}|${device.mapping}|${device.connected}`)
      .join('||');
  }

  function syncSelectedControllerId(): void {
    const selected = ctx.state.connectedGamepads[ctx.state.controllerDeviceSelectedIndex];
    selectedControllerId = selected?.id ?? null;
  }

  function syncSelectedIndexToCurrentController(): void {
    const preferredId = ctx.state.controllerConfig?.preferredGamepadId ?? '';
    const activeIndex = ctx.state.connectedGamepads.findIndex(
      (device) => device.id === ctx.state.activeGamepadId,
    );
    if (activeIndex >= 0) {
      ctx.state.controllerDeviceSelectedIndex = activeIndex;
      syncSelectedControllerId();
      return;
    }
    const preferredIndex = ctx.state.connectedGamepads.findIndex(
      (device) => device.id === preferredId,
    );
    if (preferredIndex >= 0) {
      ctx.state.controllerDeviceSelectedIndex = preferredIndex;
      syncSelectedControllerId();
      return;
    }
    clampSelectedIndex(ctx);
    syncSelectedControllerId();
  }

  function setStatus(message: string, isError = false): void {
    ctx.dom.controllerSelectStatus.textContent = message;
    ctx.dom.controllerSelectStatus.classList.toggle('error', isError);
  }

  function renderPicker(): void {
    ctx.dom.controllerSelectPicker.innerHTML = '';
    clampSelectedIndex(ctx);

    const preferredId = ctx.state.controllerConfig?.preferredGamepadId ?? '';
    ctx.state.connectedGamepads.forEach((device, index) => {
      const option = document.createElement('option');
      option.value = device.id;
      option.selected = index === ctx.state.controllerDeviceSelectedIndex;
      option.textContent = `${device.id || `Gamepad ${device.index}`} (${[
        `#${device.index}`,
        device.mapping || 'unknown',
        device.id === ctx.state.activeGamepadId ? 'active' : null,
        device.id === preferredId ? 'saved' : null,
      ]
        .filter(Boolean)
        .join(', ')})`;
      ctx.dom.controllerSelectPicker.appendChild(option);
    });

    ctx.dom.controllerSelectPicker.disabled = ctx.state.connectedGamepads.length === 0;
    ctx.dom.controllerSelectSummary.textContent =
      ctx.state.connectedGamepads.length === 0
        ? 'No controller detected.'
        : `Active: ${ctx.state.activeGamepadId ?? 'none'} · Preferred: ${preferredId || 'none'}`;

    lastRenderedDevicesKey = getDevicesKey();
    lastRenderedActiveGamepadId = ctx.state.activeGamepadId;
    lastRenderedPreferredId = preferredId;
  }

  async function saveControllerConfig(update: Parameters<typeof window.electronAPI.saveControllerConfig>[0]) {
    await window.electronAPI.saveControllerConfig(update);
    if (!ctx.state.controllerConfig) return;
    if (update.preferredGamepadId !== undefined) {
      ctx.state.controllerConfig.preferredGamepadId = update.preferredGamepadId;
    }
    if (update.preferredGamepadLabel !== undefined) {
      ctx.state.controllerConfig.preferredGamepadLabel = update.preferredGamepadLabel;
    }
    if (update.bindings) {
      ctx.state.controllerConfig.bindings = {
        ...ctx.state.controllerConfig.bindings,
        ...update.bindings,
      } as typeof ctx.state.controllerConfig.bindings;
    }
  }

  async function saveBinding(
    actionId: keyof NonNullable<typeof ctx.state.controllerConfig>['bindings'],
    binding: NonNullable<NonNullable<typeof ctx.state.controllerConfig>['bindings']>[typeof actionId],
  ): Promise<void> {
    const definition = getControllerBindingDefinition(actionId);
    try {
      await saveControllerConfig({
        bindings: {
          [actionId]: binding,
        },
      });
      learningActionId = null;
      bindingCapture = null;
      controllerConfigForm.render();
      setStatus(`${definition?.label ?? actionId} updated.`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Failed to save binding: ${message}`, true);
    }
  }

  async function saveSelectedController(): Promise<void> {
    const selected = ctx.state.connectedGamepads[ctx.state.controllerDeviceSelectedIndex];
    if (!selected) {
      setStatus('No controller selected.', true);
      return;
    }

    try {
      await saveControllerConfig({
        preferredGamepadId: selected.id,
        preferredGamepadLabel: selected.id,
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Failed to save preferred controller: ${message}`, true);
      return;
    }

    syncSelectedControllerId();
    renderPicker();
    setStatus(`Saved preferred controller: ${selected.id || `Gamepad ${selected.index}`}`);
  }

  function updateDevices(): void {
    if (!ctx.state.controllerSelectModalOpen) return;
    if (selectedControllerId) {
      const preservedIndex = ctx.state.connectedGamepads.findIndex(
        (device) => device.id === selectedControllerId,
      );
      if (preservedIndex >= 0) {
        ctx.state.controllerDeviceSelectedIndex = preservedIndex;
      } else {
        syncSelectedIndexToCurrentController();
      }
    } else {
      syncSelectedIndexToCurrentController();
    }

    if (bindingCapture && learningActionId) {
      const result = bindingCapture.poll({
        axes: ctx.state.controllerRawAxes,
        buttons: ctx.state.controllerRawButtons,
      });
      if (result) {
        void saveBinding(result.actionId as keyof NonNullable<typeof ctx.state.controllerConfig>['bindings'], result.binding as never);
      }
    }

    const preferredId = ctx.state.controllerConfig?.preferredGamepadId ?? '';
    const shouldRender =
      getDevicesKey() !== lastRenderedDevicesKey ||
      ctx.state.activeGamepadId !== lastRenderedActiveGamepadId ||
      preferredId !== lastRenderedPreferredId;
    if (shouldRender) {
      renderPicker();
      controllerConfigForm.render();
    }

    if (ctx.state.connectedGamepads.length === 0 && !learningActionId) {
      setStatus('No controllers detected.');
    }
  }

  function openControllerSelectModal(): void {
    ctx.state.controllerSelectModalOpen = true;
    syncSelectedIndexToCurrentController();
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.controllerSelectModal.classList.remove('hidden');
    ctx.dom.controllerSelectModal.setAttribute('aria-hidden', 'false');
    window.focus();
    ctx.dom.overlay.focus({ preventScroll: true });
    renderPicker();
    controllerConfigForm.render();
    if (ctx.state.connectedGamepads.length === 0) {
      setStatus('No controllers detected.');
    } else {
      setStatus('Choose a controller or click Learn to remap an action.');
    }
  }

  function closeControllerSelectModal(): void {
    if (!ctx.state.controllerSelectModalOpen) return;
    learningActionId = null;
    bindingCapture = null;
    ctx.state.controllerSelectModalOpen = false;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.controllerSelectModal.classList.add('hidden');
    ctx.dom.controllerSelectModal.setAttribute('aria-hidden', 'true');
    window.electronAPI.notifyOverlayModalClosed('controller-select');
    if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
    }
  }

  function handleControllerSelectKeydown(event: KeyboardEvent): boolean {
    if (event.key === 'Escape') {
      event.preventDefault();
      if (learningActionId) {
        learningActionId = null;
        bindingCapture = null;
        controllerConfigForm.render();
        setStatus('Controller learn mode cancelled.');
        return true;
      }
      closeControllerSelectModal();
      return true;
    }

    if (event.key === 'ArrowDown' || event.key === 'j' || event.key === 'J') {
      event.preventDefault();
      if (ctx.state.connectedGamepads.length > 0) {
        ctx.state.controllerDeviceSelectedIndex = Math.min(
          ctx.state.connectedGamepads.length - 1,
          ctx.state.controllerDeviceSelectedIndex + 1,
        );
        syncSelectedControllerId();
        renderPicker();
      }
      return true;
    }

    if (event.key === 'ArrowUp' || event.key === 'k' || event.key === 'K') {
      event.preventDefault();
      if (ctx.state.connectedGamepads.length > 0) {
        ctx.state.controllerDeviceSelectedIndex = Math.max(
          0,
          ctx.state.controllerDeviceSelectedIndex - 1,
        );
        syncSelectedControllerId();
        renderPicker();
      }
      return true;
    }

    if (event.key === 'Enter' && !learningActionId) {
      event.preventDefault();
      void saveSelectedController();
      return true;
    }

    return true;
  }

  function wireDomEvents(): void {
    ctx.dom.controllerSelectClose.addEventListener('click', () => {
      closeControllerSelectModal();
    });
    ctx.dom.controllerSelectSave.addEventListener('click', () => {
      void saveSelectedController();
    });
    ctx.dom.controllerSelectPicker.addEventListener('change', () => {
      const selectedId = ctx.dom.controllerSelectPicker.value;
      const selectedIndex = ctx.state.connectedGamepads.findIndex((device) => device.id === selectedId);
      if (selectedIndex >= 0) {
        ctx.state.controllerDeviceSelectedIndex = selectedIndex;
        syncSelectedControllerId();
        renderPicker();
      }
    });
  }

  return {
    openControllerSelectModal,
    closeControllerSelectModal,
    handleControllerSelectKeydown,
    updateDevices,
    wireDomEvents,
  };
}
