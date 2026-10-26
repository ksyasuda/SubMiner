import type { ModalStateReader, RendererContext } from '../context';
import { resolveControllerConfigForGamepad } from '../controller-profile-config.js';
import { createControllerBindingCapture } from '../handlers/controller-binding-capture.js';
import { createControllerComboCapture } from '../handlers/controller-combo-capture.js';
import {
  formatControllerReviewInput,
  getControllerScreenAction,
  getDefaultControllerScreenBindings,
  type ControllerReviewInputContext,
  type ControllerScreenBindings,
  type ControllerScreenDefinition,
} from '../../shared/controller-review';
import {
  CONTROLLER_SCREENS,
  getControllerScreen,
  type ControllerScreenBindingsKey,
} from '../../shared/controller-screens';
import {
  createControllerConfigForm,
  getControllerBindingDefinition,
  getDefaultControllerBinding,
  getDefaultDpadFallback,
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
    notifyControllerDisabled: () => void;
  },
) {
  let selectedControllerKey: string | null = null;
  let lastRenderedDevicesKey = '';
  let lastRenderedActiveGamepadId: string | null = null;
  let lastRenderedPreferredId = '';
  type ControllerBindingKey = keyof NonNullable<typeof ctx.state.controllerConfig>['bindings'];
  type ControllerBindingValue = NonNullable<
    NonNullable<typeof ctx.state.controllerConfig>['bindings']
  >[ControllerBindingKey];
  let learningActionId: ControllerBindingKey | null = null;
  let dpadLearningActionId: ControllerBindingKey | null = null;
  let bindingCapture: ReturnType<typeof createControllerBindingCapture> | null = null;
  let screenLearning: {
    screen: ControllerScreenDefinition;
    actionId: string;
    capture: ReturnType<typeof createControllerComboCapture>;
    holdLabel: string | null;
  } | null = null;

  function getSelectedController() {
    return ctx.state.connectedGamepads[ctx.state.controllerDeviceSelectedIndex] ?? null;
  }

  /**
   * The controller whose profile Alt+C shows and edits. A freshly opened window cannot see
   * any pad until a button is pressed in it, so until then use the saved (preferred)
   * controller rather than the global defaults it does not actually use.
   */
  function getSelectedControllerId(): string | null {
    return getSelectedController()?.id ?? (ctx.state.controllerConfig?.preferredGamepadId || null);
  }

  function getSelectedControllerConfig() {
    const config = ctx.state.controllerConfig;
    if (!config) return null;
    return resolveControllerConfigForGamepad(config, getSelectedControllerId());
  }

  function getScreenInputContext(): ControllerReviewInputContext {
    const config = getSelectedControllerConfig();
    return {
      mapping: getSelectedController()?.mapping ?? 'standard',
      faceButtonLabels: config?.faceButtonLabels ?? 'xbox',
      buttonIndices: config?.buttonIndices ?? {
        select: 6,
        buttonSouth: 0,
        buttonEast: 1,
        buttonWest: 2,
        buttonNorth: 3,
        leftShoulder: 4,
        rightShoulder: 5,
        leftStickPress: 9,
        rightStickPress: 10,
        leftTrigger: 6,
        rightTrigger: 7,
      },
    };
  }

  function cancelScreenLearning(): void {
    screenLearning = null;
  }

  function actionLabel(screen: ControllerScreenDefinition, actionId: string): string {
    return getControllerScreenAction(screen, actionId).label;
  }

  const controllerConfigForm = createControllerConfigForm({
    container: ctx.dom.controllerConfigList,
    getControllerMapping: () => getSelectedController()?.mapping ?? null,
    getFaceButtonLabels: () => ctx.state.controllerConfig?.faceButtonLabels,
    getBindings: () =>
      getSelectedControllerConfig()?.bindings ?? {
        toggleLookup: { kind: 'button', buttonIndex: 0 },
        closeLookup: { kind: 'button', buttonIndex: 1 },
        toggleKeyboardOnlyMode: { kind: 'button', buttonIndex: 3 },
        mineCard: { kind: 'button', buttonIndex: 2 },
        quitMpv: { kind: 'button', buttonIndex: 6 },
        previousAudio: { kind: 'none' },
        nextAudio: { kind: 'button', buttonIndex: 5 },
        playCurrentAudio: { kind: 'button', buttonIndex: 4 },
        toggleMpvPause: { kind: 'button', buttonIndex: 9 },
        holdSecondarySub: { kind: 'none' },
        previousEntry: { kind: 'none' },
        nextEntry: { kind: 'none' },
        leftStickHorizontal: { kind: 'axis', axisIndex: 0, dpadFallback: 'horizontal' },
        leftStickVertical: { kind: 'axis', axisIndex: 1, dpadFallback: 'vertical' },
        rightStickHorizontal: { kind: 'axis', axisIndex: 3, dpadFallback: 'none' },
        rightStickVertical: { kind: 'axis', axisIndex: 4, dpadFallback: 'none' },
      },
    getScreenGroups: () => {
      const config = getSelectedControllerConfig();
      if (!config) return [];
      return CONTROLLER_SCREENS.map((screen) => ({
        screen,
        bindings: config[screen.configKey] as ControllerScreenBindings,
      }));
    },
    getScreenInputContext,
    getScreenLearning: () =>
      screenLearning
        ? {
            screenId: screenLearning.screen.id,
            actionId: screenLearning.actionId,
            holdLabel: screenLearning.holdLabel,
          }
        : null,
    onScreenLearn: (screenId, actionId) => {
      learningActionId = null;
      dpadLearningActionId = null;
      bindingCapture = null;
      const config = ctx.state.controllerConfig;
      const capture = createControllerComboCapture({
        triggerDeadzone: config?.triggerDeadzone ?? 0.5,
        stickDeadzone: config?.stickDeadzone ?? 0.2,
      });
      capture.arm({ axes: ctx.state.controllerRawAxes, buttons: ctx.state.controllerRawButtons });
      const screen = getControllerScreen(screenId);
      screenLearning = { screen, actionId, capture, holdLabel: null };
      controllerConfigForm.render();
      setStatus(
        `Press a button for ${actionLabel(screen, actionId)}, or hold one and press another for a combo.`,
      );
    },
    onScreenClear: (screenId, actionId) => {
      const screen = getControllerScreen(screenId);
      void saveScreenBindings(
        screen,
        { [actionId]: { kind: 'none' } },
        `${actionLabel(screen, actionId)} cleared.`,
      );
    },
    onScreenReset: (screenId, actionId) => {
      const screen = getControllerScreen(screenId);
      void saveScreenBindings(
        screen,
        { [actionId]: structuredClone(screen.defaults[actionId]!) },
        `${actionLabel(screen, actionId)} reset.`,
      );
    },
    onScreenResetAll: (screenId) => {
      const screen = getControllerScreen(screenId);
      void saveScreenBindings(
        screen,
        getDefaultControllerScreenBindings(screen),
        `${screen.title} layout reset.`,
      );
    },
    getLearningActionId: () => learningActionId,
    getDpadLearningActionId: () => dpadLearningActionId,
    onLearn: (actionId, bindingType) => {
      const definition = getControllerBindingDefinition(actionId);
      if (!definition) return;
      dpadLearningActionId = null;
      cancelScreenLearning();
      const config = ctx.state.controllerConfig;
      bindingCapture = createControllerBindingCapture({
        triggerDeadzone: config?.triggerDeadzone ?? 0.5,
        stickDeadzone: config?.stickDeadzone ?? 0.2,
      });
      const currentBinding = getSelectedControllerConfig()?.bindings[actionId];
      const currentDpadFallback =
        currentBinding && currentBinding.kind === 'axis' && 'dpadFallback' in currentBinding
          ? currentBinding.dpadFallback
          : 'none';
      bindingCapture.arm(
        bindingType === 'axis'
          ? {
              actionId,
              bindingType: 'axis',
              dpadFallback: currentDpadFallback,
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
    onDpadLearn: (actionId) => {
      const definition = getControllerBindingDefinition(actionId);
      if (!definition) return;
      learningActionId = null;
      cancelScreenLearning();
      const config = ctx.state.controllerConfig;
      bindingCapture = createControllerBindingCapture({
        triggerDeadzone: config?.triggerDeadzone ?? 0.5,
        stickDeadzone: config?.stickDeadzone ?? 0.2,
      });
      bindingCapture.arm(
        { actionId, bindingType: 'dpad' },
        {
          axes: ctx.state.controllerRawAxes,
          buttons: ctx.state.controllerRawButtons,
        },
      );
      dpadLearningActionId = actionId;
      controllerConfigForm.render();
      setStatus(`Press a D-pad direction for ${definition.label}.`);
    },
    onDpadClear: (actionId) => {
      void saveDpadFallback(actionId, 'none');
    },
    onDpadReset: (actionId) => {
      void saveDpadFallback(actionId, getDefaultDpadFallback(actionId));
    },
  });

  function getDevicesKey(): string {
    return ctx.state.connectedGamepads
      .map((device) => `${device.id}|${device.index}|${device.mapping}|${device.connected}`)
      .join('||');
  }

  function getDeviceSelectionKey(device: { id: string; index: number }): string {
    return `${device.id}:${device.index}`;
  }

  function syncSelectedControllerId(): void {
    const selected = ctx.state.connectedGamepads[ctx.state.controllerDeviceSelectedIndex];
    selectedControllerKey = selected ? getDeviceSelectionKey(selected) : null;
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
      option.value = getDeviceSelectionKey(device);
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

  async function saveControllerConfig(
    update: Parameters<typeof window.electronAPI.saveControllerConfig>[0],
  ) {
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
    const config = ctx.state.controllerConfig;
    for (const { configKey } of CONTROLLER_SCREENS) {
      if (update[configKey]) {
        (config as Record<ControllerScreenBindingsKey, unknown>)[configKey] = {
          ...config[configKey],
          ...update[configKey],
        };
      }
    }
    if (update.profiles) {
      ctx.state.controllerConfig.profiles = ctx.state.controllerConfig.profiles ?? {};
      for (const [profileId, profileUpdate] of Object.entries(update.profiles)) {
        const currentProfile = ctx.state.controllerConfig.profiles[profileId];
        const baseProfile = currentProfile ?? {
          label: profileUpdate.label ?? profileId,
          buttonIndices: ctx.state.controllerConfig.buttonIndices,
          bindings: ctx.state.controllerConfig.bindings,
          reviewBindings: ctx.state.controllerConfig.reviewBindings,
          fieldGroupingBindings: ctx.state.controllerConfig.fieldGroupingBindings,
        };
        ctx.state.controllerConfig.profiles[profileId] = {
          label: profileUpdate.label ?? baseProfile.label,
          buttonIndices: {
            ...baseProfile.buttonIndices,
            ...(profileUpdate.buttonIndices ?? {}),
          },
          bindings: {
            ...baseProfile.bindings,
            ...(profileUpdate.bindings ?? {}),
          },
          reviewBindings: {
            ...(baseProfile.reviewBindings ?? ctx.state.controllerConfig.reviewBindings),
            ...(profileUpdate.reviewBindings ?? {}),
          },
          fieldGroupingBindings: {
            ...(baseProfile.fieldGroupingBindings ??
              ctx.state.controllerConfig.fieldGroupingBindings),
            ...(profileUpdate.fieldGroupingBindings ?? {}),
          },
        } as (typeof ctx.state.controllerConfig.profiles)[string];
      }
    }
  }

  function buildBindingConfigUpdate(
    actionId: ControllerBindingKey,
    binding: ControllerBindingValue,
  ): Parameters<typeof window.electronAPI.saveControllerConfig>[0] {
    const profileId = getSelectedControllerId();
    if (!profileId) {
      return {
        bindings: {
          [actionId]: binding,
        },
      };
    }
    return {
      profiles: {
        [profileId]: {
          bindings: {
            [actionId]: binding,
          },
        },
      },
    };
  }

  async function saveBinding(
    actionId: ControllerBindingKey,
    binding: ControllerBindingValue,
  ): Promise<void> {
    const definition = getControllerBindingDefinition(actionId);
    try {
      await saveControllerConfig(buildBindingConfigUpdate(actionId, binding));
      learningActionId = null;
      dpadLearningActionId = null;
      bindingCapture = null;
      controllerConfigForm.render();
      setStatus(`${definition?.label ?? actionId} updated.`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Failed to save binding: ${message}`, true);
    }
  }

  async function saveScreenBindings(
    screen: ControllerScreenDefinition,
    bindings: Partial<ControllerScreenBindings>,
    message: string,
  ): Promise<void> {
    const profileId = getSelectedControllerId();
    const patch = { [screen.configKey]: bindings };
    try {
      await saveControllerConfig(profileId ? { profiles: { [profileId]: patch } } : patch);
      cancelScreenLearning();
      controllerConfigForm.render();
      setStatus(message);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Failed to save ${screen.title} binding: ${message}`, true);
    }
  }

  async function saveDpadFallback(
    actionId: ControllerBindingKey,
    dpadFallback: import('../../types').ControllerDpadFallback,
  ): Promise<void> {
    const definition = getControllerBindingDefinition(actionId);
    const currentBinding = getSelectedControllerConfig()?.bindings[actionId];
    if (!currentBinding || currentBinding.kind !== 'axis') return;
    const updated = { ...currentBinding, dpadFallback };
    try {
      await saveControllerConfig(buildBindingConfigUpdate(actionId, updated));
      dpadLearningActionId = null;
      bindingCapture = null;
      controllerConfigForm.render();
      setStatus(`${definition?.label ?? actionId} D-pad updated.`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Failed to save D-pad binding: ${message}`, true);
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
    if (selectedControllerKey) {
      const preservedIndex = ctx.state.connectedGamepads.findIndex(
        (device) => getDeviceSelectionKey(device) === selectedControllerKey,
      );
      if (preservedIndex >= 0) {
        ctx.state.controllerDeviceSelectedIndex = preservedIndex;
      } else {
        syncSelectedIndexToCurrentController();
      }
    } else {
      syncSelectedIndexToCurrentController();
    }

    if (bindingCapture && (learningActionId || dpadLearningActionId)) {
      const result = bindingCapture.poll({
        axes: ctx.state.controllerRawAxes,
        buttons: ctx.state.controllerRawButtons,
      });
      if (result) {
        if (result.bindingType === 'dpad') {
          void saveDpadFallback(result.actionId as ControllerBindingKey, result.dpadDirection);
        } else {
          void saveBinding(
            result.actionId as ControllerBindingKey,
            result.binding as ControllerBindingValue,
          );
        }
      }
    }

    if (screenLearning) {
      const learning = screenLearning;
      const state = learning.capture.poll({
        axes: ctx.state.controllerRawAxes,
        buttons: ctx.state.controllerRawButtons,
      });
      if (state.status === 'holding') {
        const holdLabel = formatControllerReviewInput(state.hold, getScreenInputContext(), {
          dpadPrefix: true,
        });
        if (holdLabel !== learning.holdLabel) {
          learning.holdLabel = holdLabel;
          controllerConfigForm.render();
        }
      } else if (state.status === 'done') {
        screenLearning = null;
        void saveScreenBindings(
          learning.screen,
          { [learning.actionId]: state.binding },
          `${actionLabel(learning.screen, learning.actionId)} updated.`,
        );
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

    const idle = !learningActionId && !dpadLearningActionId && !screenLearning;
    if (ctx.state.connectedGamepads.length === 0 && idle) {
      setStatus('No controllers detected.');
    } else if (idle && ctx.dom.controllerSelectStatus.textContent === 'No controllers detected.') {
      setStatus('Choose a controller or click Learn to remap an action.');
    }
  }

  function openControllerSelectModal(): boolean {
    if (ctx.state.controllerConfig?.enabled !== true) {
      options.notifyControllerDisabled();
      return false;
    }
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
      setStatus(
        ctx.state.controllerConfig?.preferredGamepadId
          ? 'Press any button on your controller so this window can see it.'
          : 'No controllers detected.',
      );
    } else {
      setStatus('Choose a controller or click Learn to remap an action.');
    }
    return true;
  }

  function closeControllerSelectModal(): void {
    if (!ctx.state.controllerSelectModalOpen) return;
    learningActionId = null;
    dpadLearningActionId = null;
    bindingCapture = null;
    cancelScreenLearning();
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
      if (learningActionId || dpadLearningActionId || screenLearning) {
        learningActionId = null;
        dpadLearningActionId = null;
        bindingCapture = null;
        cancelScreenLearning();
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
        controllerConfigForm.render();
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
        controllerConfigForm.render();
      }
      return true;
    }

    if (event.key === 'Enter' && !learningActionId && !dpadLearningActionId && !screenLearning) {
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
      const selectedKey = ctx.dom.controllerSelectPicker.value;
      const selectedIndex = ctx.state.connectedGamepads.findIndex(
        (device) => getDeviceSelectionKey(device) === selectedKey,
      );
      if (selectedIndex >= 0) {
        ctx.state.controllerDeviceSelectedIndex = selectedIndex;
        syncSelectedControllerId();
        renderPicker();
        controllerConfigForm.render();
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
