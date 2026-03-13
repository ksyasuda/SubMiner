import type {
  ControllerDeviceInfo,
  ControllerRuntimeSnapshot,
  ResolvedControllerAxisBinding,
  ResolvedControllerConfig,
  ResolvedControllerDiscreteBinding,
} from '../../types';

type ControllerButtonState = {
  value: number;
  pressed?: boolean;
  touched?: boolean;
};

type GamepadLike = {
  id: string;
  index: number;
  connected: boolean;
  mapping: string;
  axes: readonly number[];
  buttons: readonly ControllerButtonState[];
};

type GamepadControllerOptions = {
  getGamepads: () => Array<GamepadLike | null>;
  getConfig: () => ResolvedControllerConfig;
  getKeyboardModeEnabled: () => boolean;
  getLookupWindowOpen: () => boolean;
  getInteractionBlocked: () => boolean;
  toggleKeyboardMode: () => void;
  toggleLookup: () => void;
  closeLookup: () => void;
  moveSelection: (delta: -1 | 1) => void;
  mineCard: () => void;
  quitMpv: () => void;
  previousAudio: () => void;
  nextAudio: () => void;
  playCurrentAudio: () => void;
  toggleMpvPause: () => void;
  scrollPopup: (deltaPixels: number) => void;
  jumpPopup: (deltaPixels: number) => void;
  onState: (state: ControllerRuntimeSnapshot) => void;
};

type HoldState = {
  repeatStarted: boolean;
  direction: -1 | 1 | null;
  lastFireAt: number;
  initialFired: boolean;
};

const DPAD_BUTTON_INDEX = {
  up: 12,
  down: 13,
  left: 14,
  right: 15,
} as const;

const DPAD_AXIS_INDEX = {
  horizontal: 6,
  vertical: 7,
} as const;

function normalizeRawButtonState(
  button: ControllerButtonState | undefined,
  triggerDeadzone: number,
): boolean {
  if (!button) return false;
  return Boolean(button.pressed) || button.value >= triggerDeadzone;
}

function resolveGamepadAxis(gamepad: GamepadLike, axisIndex: number): number {
  const value = gamepad.axes[axisIndex];
  return typeof value === 'number' && Number.isFinite(value) ? value : 0;
}

function resolveDpadValue(
  gamepad: GamepadLike,
  negativeIndex: number,
  positiveIndex: number,
  triggerDeadzone: number,
): number {
  const negative = gamepad.buttons[negativeIndex];
  const positive = gamepad.buttons[positiveIndex];
  return (
    (normalizeRawButtonState(positive, triggerDeadzone) ? 1 : 0) -
    (normalizeRawButtonState(negative, triggerDeadzone) ? 1 : 0)
  );
}

function resolveDpadAxisValue(gamepad: GamepadLike, axisIndex: number): number {
  const value = resolveGamepadAxis(gamepad, axisIndex);
  if (Math.abs(value) < 0.5) {
    return 0;
  }
  return Math.sign(value);
}

function resolveDpadHorizontalValue(gamepad: GamepadLike, triggerDeadzone: number): number {
  const axisValue = resolveDpadAxisValue(gamepad, DPAD_AXIS_INDEX.horizontal);
  if (axisValue !== 0) {
    return axisValue;
  }
  return resolveDpadValue(
    gamepad,
    DPAD_BUTTON_INDEX.left,
    DPAD_BUTTON_INDEX.right,
    triggerDeadzone,
  );
}

function resolveDpadVerticalValue(gamepad: GamepadLike, triggerDeadzone: number): number {
  const axisValue = resolveDpadAxisValue(gamepad, DPAD_AXIS_INDEX.vertical);
  if (axisValue !== 0) {
    return axisValue;
  }
  return resolveDpadValue(gamepad, DPAD_BUTTON_INDEX.up, DPAD_BUTTON_INDEX.down, triggerDeadzone);
}

function resolveConnectedGamepads(gamepads: Array<GamepadLike | null>): GamepadLike[] {
  return gamepads
    .filter((gamepad): gamepad is GamepadLike => Boolean(gamepad?.connected))
    .sort((left, right) => left.index - right.index);
}

function createHoldState(): HoldState {
  return {
    repeatStarted: false,
    direction: null,
    lastFireAt: 0,
    initialFired: false,
  };
}

function shouldFireHeldAction(
  state: HoldState,
  now: number,
  repeatDelayMs: number,
  repeatIntervalMs: number,
): boolean {
  if (!state.initialFired) {
    state.initialFired = true;
    state.lastFireAt = now;
    return true;
  }

  const elapsed = now - state.lastFireAt;
  const threshold = state.repeatStarted ? repeatIntervalMs : repeatDelayMs;
  if (elapsed < threshold) {
    return false;
  }

  state.repeatStarted = true;
  state.lastFireAt = now;
  return true;
}

function resetHeldAction(state: HoldState): void {
  state.repeatStarted = false;
  state.direction = null;
  state.lastFireAt = 0;
  state.initialFired = false;
}

function syncHeldActionBlocked(
  state: HoldState,
  value: number,
  now: number,
  activationThreshold: number,
): void {
  if (Math.abs(value) < activationThreshold) {
    resetHeldAction(state);
    return;
  }

  const direction = value > 0 ? 1 : -1;
  state.repeatStarted = false;
  state.direction = direction;
  state.lastFireAt = now;
  state.initialFired = true;
}

function resolveDiscreteBindingPressed(
  gamepad: GamepadLike,
  binding: ResolvedControllerDiscreteBinding,
  config: ResolvedControllerConfig,
): boolean {
  if (binding.kind === 'none') {
    return false;
  }

  if (binding.kind === 'button') {
    return normalizeRawButtonState(gamepad.buttons[binding.buttonIndex], config.triggerDeadzone);
  }

  const activationThreshold = Math.max(config.stickDeadzone, 0.55);
  const axisValue = resolveGamepadAxis(gamepad, binding.axisIndex);
  return binding.direction === 'positive'
    ? axisValue >= activationThreshold
    : axisValue <= -activationThreshold;
}

function resolveAxisBindingValue(
  gamepad: GamepadLike,
  binding: ResolvedControllerAxisBinding,
  triggerDeadzone: number,
  activationThreshold: number,
): number {
  if (binding.kind === 'none') {
    return 0;
  }
  const axisValue = resolveGamepadAxis(gamepad, binding.axisIndex);
  if (Math.abs(axisValue) >= activationThreshold) {
    return axisValue;
  }

  if (binding.dpadFallback === 'horizontal') {
    return resolveDpadHorizontalValue(gamepad, triggerDeadzone);
  }
  if (binding.dpadFallback === 'vertical') {
    return resolveDpadVerticalValue(gamepad, triggerDeadzone);
  }
  return axisValue;
}

export function createGamepadController(options: GamepadControllerOptions) {
  let previousActions = new Map<string, boolean>();
  let selectionHold = createHoldState();
  let jumpHold = createHoldState();
  let activeGamepadId: string | null = null;
  let lastPollAt: number | null = null;

  function getConnectedGamepads(): GamepadLike[] {
    return resolveConnectedGamepads(options.getGamepads());
  }

  function resolveActiveGamepad(
    gamepads: GamepadLike[],
    config: ResolvedControllerConfig,
  ): GamepadLike | null {
    if (gamepads.length === 0) return null;
    if (config.preferredGamepadId.trim().length > 0) {
      const preferred = gamepads.find((gamepad) => gamepad.id === config.preferredGamepadId);
      if (preferred) {
        return preferred;
      }
    }
    return gamepads[0] ?? null;
  }

  function publishState(gamepads: GamepadLike[], activeGamepad: GamepadLike | null): void {
    activeGamepadId = activeGamepad?.id ?? null;
    options.onState({
      connectedGamepads: gamepads.map((gamepad) => ({
        id: gamepad.id,
        index: gamepad.index,
        mapping: gamepad.mapping,
        connected: gamepad.connected,
      })) satisfies ControllerDeviceInfo[],
      activeGamepadId,
      rawAxes: activeGamepad?.axes ? [...activeGamepad.axes] : [],
      rawButtons: activeGamepad?.buttons
        ? activeGamepad.buttons.map((button) => ({
            value: button.value,
            pressed: Boolean(button.pressed),
            touched: button.touched,
          }))
        : [],
    });
  }

  function handleActionEdge(
    actionKey: string,
    binding: ResolvedControllerDiscreteBinding,
    activeGamepad: GamepadLike,
    config: ResolvedControllerConfig,
    action: () => void,
  ): void {
    const isPressed = resolveDiscreteBindingPressed(activeGamepad, binding, config);
    const wasPressed = previousActions.get(actionKey) ?? false;
    previousActions.set(actionKey, isPressed);
    if (!wasPressed && isPressed) {
      action();
    }
  }

  function handleSelectionAxis(value: number, now: number, config: ResolvedControllerConfig): void {
    const activationThreshold = Math.max(config.stickDeadzone, 0.55);
    if (Math.abs(value) < activationThreshold) {
      resetHeldAction(selectionHold);
      return;
    }

    const direction = value > 0 ? 1 : -1;
    if (selectionHold.direction !== direction) {
      resetHeldAction(selectionHold);
      selectionHold.direction = direction;
    }

    if (shouldFireHeldAction(selectionHold, now, config.repeatDelayMs, config.repeatIntervalMs)) {
      options.moveSelection(direction);
    }
  }

  function handleJumpAxis(value: number, now: number, config: ResolvedControllerConfig): void {
    const activationThreshold = Math.max(config.stickDeadzone, 0.55);
    if (Math.abs(value) < activationThreshold) {
      resetHeldAction(jumpHold);
      return;
    }

    const direction = value > 0 ? 1 : -1;
    if (jumpHold.direction !== direction) {
      resetHeldAction(jumpHold);
      jumpHold.direction = direction;
    }

    if (shouldFireHeldAction(jumpHold, now, config.repeatDelayMs, config.repeatIntervalMs)) {
      options.jumpPopup(direction * config.horizontalJumpPixels);
    }
  }

  function syncBlockedInteractionState(
    activeGamepad: GamepadLike,
    config: ResolvedControllerConfig,
    now: number,
  ): void {
    const discreteActions = [
      ['toggleKeyboardOnlyMode', config.bindings.toggleKeyboardOnlyMode],
      ['toggleLookup', config.bindings.toggleLookup],
      ['closeLookup', config.bindings.closeLookup],
      ['mineCard', config.bindings.mineCard],
      ['quitMpv', config.bindings.quitMpv],
      ['previousAudio', config.bindings.previousAudio],
      ['nextAudio', config.bindings.nextAudio],
      ['playCurrentAudio', config.bindings.playCurrentAudio],
      ['toggleMpvPause', config.bindings.toggleMpvPause],
    ] as const;

    for (const [actionKey, binding] of discreteActions) {
      previousActions.set(actionKey, resolveDiscreteBindingPressed(activeGamepad, binding, config));
    }

    const activationThreshold = Math.max(config.stickDeadzone, 0.55);
    const selectionValue = resolveAxisBindingValue(
      activeGamepad,
      config.bindings.leftStickHorizontal,
      config.triggerDeadzone,
      activationThreshold,
    );
    syncHeldActionBlocked(selectionHold, selectionValue, now, activationThreshold);

    if (options.getLookupWindowOpen()) {
      syncHeldActionBlocked(
        jumpHold,
        resolveAxisBindingValue(
          activeGamepad,
          config.bindings.rightStickVertical,
          config.triggerDeadzone,
          activationThreshold,
        ),
        now,
        activationThreshold,
      );
    } else {
      resetHeldAction(jumpHold);
    }
  }

  function poll(now: number): void {
    const elapsedMs = lastPollAt === null ? 0 : Math.max(now - lastPollAt, 0);
    lastPollAt = now;
    const config = options.getConfig();
    const connectedGamepads = getConnectedGamepads();
    const activeGamepad = resolveActiveGamepad(connectedGamepads, config);
    const previousActiveGamepadId = activeGamepadId;
    publishState(connectedGamepads, activeGamepad);

    if (!activeGamepad) {
      previousActions = new Map();
      resetHeldAction(selectionHold);
      resetHeldAction(jumpHold);
      lastPollAt = null;
      return;
    }

    if (activeGamepad.id !== previousActiveGamepadId) {
      previousActions = new Map();
      resetHeldAction(selectionHold);
      resetHeldAction(jumpHold);
    }

    let interactionAllowed =
      config.enabled && options.getKeyboardModeEnabled() && !options.getInteractionBlocked();
    if (config.enabled) {
      handleActionEdge(
        'toggleKeyboardOnlyMode',
        config.bindings.toggleKeyboardOnlyMode,
        activeGamepad,
        config,
        options.toggleKeyboardMode,
      );
    }

    interactionAllowed =
      config.enabled && options.getKeyboardModeEnabled() && !options.getInteractionBlocked();

    if (!interactionAllowed) {
      syncBlockedInteractionState(activeGamepad, config, now);
      return;
    }

    handleActionEdge(
      'toggleLookup',
      config.bindings.toggleLookup,
      activeGamepad,
      config,
      options.toggleLookup,
    );
    handleActionEdge(
      'closeLookup',
      config.bindings.closeLookup,
      activeGamepad,
      config,
      options.closeLookup,
    );
    handleActionEdge('mineCard', config.bindings.mineCard, activeGamepad, config, options.mineCard);
    handleActionEdge('quitMpv', config.bindings.quitMpv, activeGamepad, config, options.quitMpv);

    const activationThreshold = Math.max(config.stickDeadzone, 0.55);

    if (options.getLookupWindowOpen()) {
      handleActionEdge(
        'previousAudio',
        config.bindings.previousAudio,
        activeGamepad,
        config,
        options.previousAudio,
      );
      handleActionEdge(
        'nextAudio',
        config.bindings.nextAudio,
        activeGamepad,
        config,
        options.nextAudio,
      );
      handleActionEdge(
        'playCurrentAudio',
        config.bindings.playCurrentAudio,
        activeGamepad,
        config,
        options.playCurrentAudio,
      );

      const primaryScroll = resolveAxisBindingValue(
        activeGamepad,
        config.bindings.leftStickVertical,
        config.triggerDeadzone,
        config.stickDeadzone,
      );
      if (elapsedMs > 0 && Math.abs(primaryScroll) >= config.stickDeadzone) {
        options.scrollPopup((primaryScroll * config.scrollPixelsPerSecond * elapsedMs) / 1000);
      }

      handleJumpAxis(
        resolveAxisBindingValue(
          activeGamepad,
          config.bindings.rightStickVertical,
          config.triggerDeadzone,
          activationThreshold,
        ),
        now,
        config,
      );
    } else {
      resetHeldAction(jumpHold);
    }

    handleActionEdge(
      'toggleMpvPause',
      config.bindings.toggleMpvPause,
      activeGamepad,
      config,
      options.toggleMpvPause,
    );

    handleSelectionAxis(
      resolveAxisBindingValue(
        activeGamepad,
        config.bindings.leftStickHorizontal,
        config.triggerDeadzone,
        activationThreshold,
      ),
      now,
      config,
    );
  }

  return {
    poll,
    getActiveGamepadId: (): string | null => activeGamepadId,
  };
}
