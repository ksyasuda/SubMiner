import type {
  ControllerAxisBinding,
  ControllerButtonBinding,
  ControllerDeviceInfo,
  ControllerRuntimeSnapshot,
  ControllerTriggerInputMode,
  ResolvedControllerConfig,
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

const DEFAULT_BUTTON_INDEX_BY_BINDING: Record<Exclude<ControllerButtonBinding, 'none'>, number> = {
  select: 8,
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
};

const AXIS_INDEX_BY_BINDING: Record<ControllerAxisBinding, number> = {
  leftStickX: 0,
  leftStickY: 1,
  rightStickX: 3,
  rightStickY: 4,
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

function isTriggerBinding(binding: ControllerButtonBinding): boolean {
  return binding === 'leftTrigger' || binding === 'rightTrigger';
}

function resolveButtonIndex(
  config: ResolvedControllerConfig,
  binding: ControllerButtonBinding,
): number {
  if (binding === 'none') {
    return -1;
  }
  return config.buttonIndices[binding] ?? DEFAULT_BUTTON_INDEX_BY_BINDING[binding];
}

function normalizeButtonState(
  gamepad: GamepadLike,
  config: ResolvedControllerConfig,
  binding: ControllerButtonBinding,
  triggerInputMode: ControllerTriggerInputMode,
  triggerDeadzone: number,
): boolean {
  if (binding === 'none') {
    return false;
  }
  const button = gamepad.buttons[resolveButtonIndex(config, binding)];
  if (isTriggerBinding(binding)) {
    return normalizeTriggerState(button, triggerInputMode, triggerDeadzone);
  }
  return normalizeRawButtonState(button, triggerDeadzone);
}

function normalizeRawButtonState(
  button: ControllerButtonState | undefined,
  triggerDeadzone: number,
): boolean {
  if (!button) return false;
  return Boolean(button.pressed) || button.value >= triggerDeadzone;
}

function normalizeTriggerState(
  button: ControllerButtonState | undefined,
  mode: ControllerTriggerInputMode,
  triggerDeadzone: number,
): boolean {
  if (!button) return false;
  if (mode === 'digital') {
    return Boolean(button.pressed);
  }
  if (mode === 'analog') {
    return button.value >= triggerDeadzone;
  }
  return Boolean(button.pressed) || button.value >= triggerDeadzone;
}

function resolveAxisValue(gamepad: GamepadLike, binding: ControllerAxisBinding): number {
  return gamepad.axes[AXIS_INDEX_BY_BINDING[binding]] ?? 0;
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

function resolveDpadAxisValue(
  gamepad: GamepadLike,
  axisIndex: number,
): number {
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
  return resolveDpadValue(gamepad, DPAD_BUTTON_INDEX.left, DPAD_BUTTON_INDEX.right, triggerDeadzone);
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

function shouldFireHeldAction(state: HoldState, now: number, repeatDelayMs: number, repeatIntervalMs: number): boolean {
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

export function createGamepadController(options: GamepadControllerOptions) {
  let previousButtons = new Map<ControllerButtonBinding, boolean>();
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

  function handleButtonEdge(
    binding: ControllerButtonBinding,
    isPressed: boolean,
    action: () => void,
  ): void {
    if (binding === 'none') {
      return;
    }
    const wasPressed = previousButtons.get(binding) ?? false;
    previousButtons.set(binding, isPressed);
    if (!wasPressed && isPressed) {
      action();
    }
  }

  function handleSelectionAxis(
    value: number,
    now: number,
    config: ResolvedControllerConfig,
  ): void {
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

  function handleJumpAxis(
    value: number,
    now: number,
    config: ResolvedControllerConfig,
  ): void {
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

  function poll(now: number): void {
    const elapsedMs = lastPollAt === null ? 0 : Math.max(now - lastPollAt, 0);
    lastPollAt = now;
    const config = options.getConfig();
    const connectedGamepads = getConnectedGamepads();
    const activeGamepad = resolveActiveGamepad(connectedGamepads, config);
    publishState(connectedGamepads, activeGamepad);

    if (!activeGamepad) {
      previousButtons = new Map();
      resetHeldAction(selectionHold);
      resetHeldAction(jumpHold);
      lastPollAt = null;
      return;
    }

    handleButtonEdge(
      config.bindings.toggleKeyboardOnlyMode,
      normalizeButtonState(
        activeGamepad,
        config,
        config.bindings.toggleKeyboardOnlyMode,
        config.triggerInputMode,
        config.triggerDeadzone,
      ),
      options.toggleKeyboardMode,
    );

    const interactionAllowed =
      config.enabled &&
      options.getKeyboardModeEnabled() &&
      !options.getInteractionBlocked();
    if (!interactionAllowed) {
      return;
    }

    handleButtonEdge(
      config.bindings.toggleLookup,
      normalizeButtonState(
        activeGamepad,
        config,
        config.bindings.toggleLookup,
        config.triggerInputMode,
        config.triggerDeadzone,
      ),
      options.toggleLookup,
    );
    handleButtonEdge(
      config.bindings.closeLookup,
      normalizeButtonState(
        activeGamepad,
        config,
        config.bindings.closeLookup,
        config.triggerInputMode,
        config.triggerDeadzone,
      ),
      options.closeLookup,
    );
    handleButtonEdge(
      config.bindings.mineCard,
      normalizeButtonState(
        activeGamepad,
        config,
        config.bindings.mineCard,
        config.triggerInputMode,
        config.triggerDeadzone,
      ),
      options.mineCard,
    );
    handleButtonEdge(
      config.bindings.quitMpv,
      normalizeButtonState(
        activeGamepad,
        config,
        config.bindings.quitMpv,
        config.triggerInputMode,
        config.triggerDeadzone,
      ),
      options.quitMpv,
    );

    if (options.getLookupWindowOpen()) {
      handleButtonEdge(
        config.bindings.previousAudio,
        normalizeButtonState(
          activeGamepad,
          config,
          config.bindings.previousAudio,
          config.triggerInputMode,
          config.triggerDeadzone,
        ),
        options.previousAudio,
      );
      handleButtonEdge(
        config.bindings.nextAudio,
        normalizeButtonState(
          activeGamepad,
          config,
          config.bindings.nextAudio,
          config.triggerInputMode,
          config.triggerDeadzone,
        ),
        options.nextAudio,
      );
      handleButtonEdge(
        config.bindings.playCurrentAudio,
        normalizeButtonState(
          activeGamepad,
          config,
          config.bindings.playCurrentAudio,
          config.triggerInputMode,
          config.triggerDeadzone,
        ),
        options.playCurrentAudio,
      );

      const dpadVertical = resolveDpadVerticalValue(activeGamepad, config.triggerDeadzone);
      const primaryScroll = resolveAxisValue(activeGamepad, config.bindings.leftStickVertical);
      if (elapsedMs > 0) {
        if (Math.abs(primaryScroll) >= config.stickDeadzone) {
          options.scrollPopup((primaryScroll * config.scrollPixelsPerSecond * elapsedMs) / 1000);
        }
        if (dpadVertical !== 0) {
          options.scrollPopup((dpadVertical * config.scrollPixelsPerSecond * elapsedMs) / 1000);
        }
      }

      handleJumpAxis(
        resolveAxisValue(activeGamepad, config.bindings.rightStickVertical),
        now,
        config,
      );
    } else {
      resetHeldAction(jumpHold);
    }

    handleButtonEdge(
      config.bindings.toggleMpvPause,
      normalizeButtonState(
        activeGamepad,
        config,
        config.bindings.toggleMpvPause,
        config.triggerInputMode,
        config.triggerDeadzone,
      ),
      options.toggleMpvPause,
    );

    handleSelectionAxis(
      (() => {
        const axisValue = resolveAxisValue(activeGamepad, config.bindings.leftStickHorizontal);
        if (Math.abs(axisValue) >= Math.max(config.stickDeadzone, 0.55)) {
          return axisValue;
        }
        return resolveDpadHorizontalValue(activeGamepad, config.triggerDeadzone);
      })(),
      now,
      config,
    );
  }

  return {
    poll,
    getActiveGamepadId: (): string | null => activeGamepadId,
  };
}
