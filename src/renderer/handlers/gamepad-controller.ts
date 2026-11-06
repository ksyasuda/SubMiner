import type {
  ControllerDeviceInfo,
  ControllerRuntimeSnapshot,
  ResolvedControllerAxisBinding,
  ResolvedControllerConfig,
  ResolvedControllerDiscreteBinding,
} from '../../types';
import { resolveControllerConfigForGamepad } from '../controller-profile-config.js';
import { createControllerComboTracker } from './controller-combo-tracker.js';

type ControllerButtonState = {
  value: number;
  pressed?: boolean;
  touched?: boolean;
};

export type GamepadLike = {
  id: string;
  index: number;
  connected: boolean;
  mapping: string;
  /** When the pad last reported new input (Gamepad API). */
  timestamp?: number;
  axes: readonly number[];
  buttons: readonly ControllerButtonState[];
};

/**
 * A screen that takes over every controller input while it is open (for example the
 * media timing review). `activate` runs on the first poll of a new session; with
 * `latchHeld` it ignores buttons already held (so the press that opened the screen does
 * not fire), otherwise it treats them as fresh presses. `poll` runs on every later frame.
 */
export type ControllerScreenLayout = {
  activate: (
    gamepad: GamepadLike,
    config: ResolvedControllerConfig,
    now: number,
    options: { latchHeld: boolean },
  ) => void;
  poll: (gamepad: GamepadLike, config: ResolvedControllerConfig, now: number) => void;
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
  previousEntry: () => void;
  nextEntry: () => void;
  /** Switches keyboard-only mode between moving by word and by character. */
  toggleCharacterNavigation: () => void;
  toggleMpvPause: () => void;
  previousSubtitle: () => void;
  nextSubtitle: () => void;
  replaySubtitle: () => void;
  playNextSubtitle: () => void;
  setSecondarySubHeld: (held: boolean) => void;
  scrollPopup: (deltaPixels: number) => void;
  jumpPopup: (deltaPixels: number) => void;
  onState: (state: ControllerRuntimeSnapshot) => void;
  /** Screen layout that currently owns the controller in this window, if any. */
  getScreenLayout?: () => ControllerScreenLayout | null;
  /** True while a modal in another window owns the controller; this window stands down. */
  getSuspended?: () => boolean;
};

export type HoldState = {
  repeatStarted: boolean;
  direction: -1 | 1 | null;
  lastFireAt: number;
  initialFired: boolean;
};

/** Buttons read as the D-pad on every pad, whatever its reported mapping. */
export const DPAD_BUTTON_INDEX = {
  up: 12,
  down: 13,
  left: 14,
  right: 15,
} as const;

/** Axes some raw (non-standard) pads use for the D-pad. */
export const DPAD_AXIS_INDEX = {
  horizontal: 6,
  vertical: 7,
} as const;

export function normalizeRawButtonState(
  button: ControllerButtonState | undefined,
  triggerDeadzone: number,
): boolean {
  if (!button) return false;
  return Boolean(button.pressed) || button.value >= triggerDeadzone;
}

function resolveTriggerBindingPressed(
  button: ControllerButtonState | undefined,
  config: ResolvedControllerConfig,
): boolean {
  if (!button) return false;
  if (config.triggerInputMode === 'digital') {
    return Boolean(button.pressed);
  }
  if (config.triggerInputMode === 'analog') {
    return button.value >= config.triggerDeadzone;
  }
  return normalizeRawButtonState(button, config.triggerDeadzone);
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

export function resolveDpadHorizontalValue(gamepad: GamepadLike, triggerDeadzone: number): number {
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

export function resolveDpadVerticalValue(gamepad: GamepadLike, triggerDeadzone: number): number {
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

export function createHoldState(): HoldState {
  return {
    repeatStarted: false,
    direction: null,
    lastFireAt: 0,
    initialFired: false,
  };
}

export function shouldFireHeldAction(
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

export function resetHeldAction(state: HoldState): void {
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

/** Whether the binding's press input is down. A combo's `hold` is checked by the combo tracker. */
export function resolveDiscreteBindingPressed(
  gamepad: GamepadLike,
  binding: ResolvedControllerDiscreteBinding,
  config: ResolvedControllerConfig,
): boolean {
  if (binding.kind === 'none') {
    return false;
  }

  if (binding.kind === 'button') {
    const button = gamepad.buttons[binding.buttonIndex];
    const isTriggerBinding =
      binding.buttonIndex === config.buttonIndices.leftTrigger ||
      binding.buttonIndex === config.buttonIndices.rightTrigger;
    return isTriggerBinding
      ? resolveTriggerBindingPressed(button, config)
      : normalizeRawButtonState(button, config.triggerDeadzone);
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

/** Discrete actions that only act while the lookup popup is open. */
const POPUP_ONLY_ACTIONS = [
  'previousAudio',
  'nextAudio',
  'playCurrentAudio',
  'previousEntry',
  'nextEntry',
] as const;

/** Playback and subtitle actions; like the secondary hold, they do not need keyboard-only mode. */
const WATCHING_ACTIONS = [
  'toggleMpvPause',
  'previousSubtitle',
  'nextSubtitle',
  'replaySubtitle',
  'playNextSubtitle',
] as const;

/** Every action fired by a fresh press; `holdSecondarySub` follows the button instead. */
const PRESS_ACTIONS = [
  'toggleKeyboardOnlyMode',
  'toggleLookup',
  'closeLookup',
  'mineCard',
  'quitMpv',
  'toggleCharacterNavigation',
  ...POPUP_ONLY_ACTIONS,
  ...WATCHING_ACTIONS,
] as const;

export function createGamepadController(options: GamepadControllerOptions) {
  let previousActions = new Map<string, boolean>();
  const comboTracker = createControllerComboTracker();
  /** Plain actions on a hold button, pressed and waiting to see whether they were a tap. */
  const pendingTaps = new Set<string>();
  const handledThisPoll = new Set<string>();
  let selectionHold = createHoldState();
  let jumpHold = createHoldState();
  let activeGamepadId: string | null = null;
  let lastPollAt: number | null = null;
  let secondarySubHeld = false;
  let activeScreenLayout: ControllerScreenLayout | null = null;
  let screenLayoutWaitingForPad = false;

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
        timestamp: gamepad.timestamp,
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

  // Level-triggered: reports every change in held state, so releasing the button,
  // disconnecting the pad, disabling the controller, or opening a blocking modal
  // all hide the secondary subtitle again instead of leaving it stuck visible.
  function syncSecondarySubHold(held: boolean): void {
    if (held === secondarySubHeld) return;
    secondarySubHeld = held;
    options.setSecondarySubHeld(held);
  }

  function handleActionEdge(
    actionKey: string,
    binding: ResolvedControllerDiscreteBinding,
    activeGamepad: GamepadLike,
    config: ResolvedControllerConfig,
    action: () => void,
  ): void {
    handledThisPoll.add(actionKey);
    const isPressed = resolveDiscreteBindingPressed(activeGamepad, binding, config);
    const wasPressed = previousActions.get(actionKey) ?? false;
    previousActions.set(actionKey, isPressed);
    if (isPressed === wasPressed) return;

    if (!isPressed) {
      if (pendingTaps.delete(actionKey) && !comboTracker.isTapCancelled(binding)) action();
      return;
    }
    if (!comboTracker.matches(binding)) return;
    if (comboTracker.firesOnRelease(binding)) {
      pendingTaps.add(actionKey);
    } else {
      action();
    }
  }

  // The action does not apply this poll. Its edge state keeps following the pad, so a release
  // is not missed and a held button waits for a fresh press, and a pending tap is dropped.
  function trackActionWithoutFiring(
    actionKey: (typeof PRESS_ACTIONS)[number],
    activeGamepad: GamepadLike,
    config: ResolvedControllerConfig,
  ): void {
    previousActions.set(
      actionKey,
      resolveDiscreteBindingPressed(activeGamepad, config.bindings[actionKey], config),
    );
    if (!handledThisPoll.has(actionKey)) pendingTaps.delete(actionKey);
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
    for (const actionKey of PRESS_ACTIONS) {
      trackActionWithoutFiring(actionKey, activeGamepad, config);
    }

    const axisGamepad = comboTracker.withoutClaimedInputs(activeGamepad);
    const activationThreshold = Math.max(config.stickDeadzone, 0.55);
    const selectionValue = resolveAxisBindingValue(
      axisGamepad,
      config.bindings.leftStickHorizontal,
      config.triggerDeadzone,
      activationThreshold,
    );
    syncHeldActionBlocked(selectionHold, selectionValue, now, activationThreshold);

    if (options.getLookupWindowOpen()) {
      syncHeldActionBlocked(
        jumpHold,
        resolveAxisBindingValue(
          axisGamepad,
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
    handledThisPoll.clear();

    if (!activeGamepad) {
      syncSecondarySubHold(false);
      activeScreenLayout = null;
      // Browsers hide a pad from a page until a button is pressed while it is open, and
      // each review opens in a fresh window. When a screen is already waiting for the pad,
      // the press that reveals it started after the screen opened, so it should count.
      screenLayoutWaitingForPad = config.enabled && (options.getScreenLayout?.() ?? null) !== null;
      previousActions = new Map();
      comboTracker.reset();
      pendingTaps.clear();
      resetHeldAction(selectionHold);
      resetHeldAction(jumpHold);
      lastPollAt = null;
      return;
    }

    if (activeGamepad.id !== previousActiveGamepadId) {
      activeScreenLayout = null;
      previousActions = new Map();
      comboTracker.reset();
      pendingTaps.clear();
      resetHeldAction(selectionHold);
      resetHeldAction(jumpHold);
    }

    const activeConfig = resolveControllerConfigForGamepad(config, activeGamepad.id);
    comboTracker.update(
      [...PRESS_ACTIONS, 'holdSecondarySub' as const].map((key) => activeConfig.bindings[key]),
      (input) => resolveDiscreteBindingPressed(activeGamepad, input, activeConfig),
    );

    // A screen layout or another window owning the controller pauses every normal action,
    // including the keyboard-mode toggle and the secondary subtitle hold. Their edge state
    // keeps tracking the pad, so a button still held when control returns does not fire.
    const screenLayout = activeConfig.enabled ? (options.getScreenLayout?.() ?? null) : null;
    const suspended = screenLayout === null && (options.getSuspended?.() ?? false);
    if (screenLayout !== null || suspended) {
      syncSecondarySubHold(false);
      syncBlockedInteractionState(activeGamepad, activeConfig, now);
      if (screenLayout === null) {
        activeScreenLayout = null;
        screenLayoutWaitingForPad = false;
      } else if (screenLayout !== activeScreenLayout) {
        activeScreenLayout = screenLayout;
        screenLayout.activate(activeGamepad, activeConfig, now, {
          latchHeld: !screenLayoutWaitingForPad,
        });
        screenLayoutWaitingForPad = false;
      } else {
        screenLayout.poll(activeGamepad, activeConfig, now);
      }
      return;
    }
    activeScreenLayout = null;
    screenLayoutWaitingForPad = false;

    // Unlike most actions, this one does not require keyboard-only mode, so it can be
    // used while simply watching.
    syncSecondarySubHold(
      activeConfig.enabled &&
        !options.getInteractionBlocked() &&
        resolveDiscreteBindingPressed(
          activeGamepad,
          activeConfig.bindings.holdSecondarySub,
          activeConfig,
        ) &&
        comboTracker.matches(activeConfig.bindings.holdSecondarySub),
    );

    // Pause and subtitle navigation also work while simply watching; a blocking modal pauses them.
    if (activeConfig.enabled && !options.getInteractionBlocked()) {
      for (const actionKey of WATCHING_ACTIONS) {
        handleActionEdge(
          actionKey,
          activeConfig.bindings[actionKey],
          activeGamepad,
          activeConfig,
          options[actionKey],
        );
      }
    }

    if (activeConfig.enabled) {
      handleActionEdge(
        'toggleKeyboardOnlyMode',
        activeConfig.bindings.toggleKeyboardOnlyMode,
        activeGamepad,
        activeConfig,
        options.toggleKeyboardMode,
      );
    }

    const interactionAllowed =
      activeConfig.enabled && options.getKeyboardModeEnabled() && !options.getInteractionBlocked();

    if (!interactionAllowed) {
      syncBlockedInteractionState(activeGamepad, activeConfig, now);
      return;
    }

    handleActionEdge(
      'toggleLookup',
      activeConfig.bindings.toggleLookup,
      activeGamepad,
      activeConfig,
      options.toggleLookup,
    );
    handleActionEdge(
      'closeLookup',
      activeConfig.bindings.closeLookup,
      activeGamepad,
      activeConfig,
      options.closeLookup,
    );
    handleActionEdge(
      'mineCard',
      activeConfig.bindings.mineCard,
      activeGamepad,
      activeConfig,
      options.mineCard,
    );
    handleActionEdge(
      'quitMpv',
      activeConfig.bindings.quitMpv,
      activeGamepad,
      activeConfig,
      options.quitMpv,
    );
    handleActionEdge(
      'toggleCharacterNavigation',
      activeConfig.bindings.toggleCharacterNavigation,
      activeGamepad,
      activeConfig,
      options.toggleCharacterNavigation,
    );

    const activationThreshold = Math.max(activeConfig.stickDeadzone, 0.55);
    const axisGamepad = comboTracker.withoutClaimedInputs(activeGamepad);

    if (options.getLookupWindowOpen()) {
      handleActionEdge(
        'previousAudio',
        activeConfig.bindings.previousAudio,
        activeGamepad,
        activeConfig,
        options.previousAudio,
      );
      handleActionEdge(
        'nextAudio',
        activeConfig.bindings.nextAudio,
        activeGamepad,
        activeConfig,
        options.nextAudio,
      );
      handleActionEdge(
        'playCurrentAudio',
        activeConfig.bindings.playCurrentAudio,
        activeGamepad,
        activeConfig,
        options.playCurrentAudio,
      );
      handleActionEdge(
        'previousEntry',
        activeConfig.bindings.previousEntry,
        activeGamepad,
        activeConfig,
        options.previousEntry,
      );
      handleActionEdge(
        'nextEntry',
        activeConfig.bindings.nextEntry,
        activeGamepad,
        activeConfig,
        options.nextEntry,
      );

      const primaryScroll = resolveAxisBindingValue(
        axisGamepad,
        activeConfig.bindings.leftStickVertical,
        activeConfig.triggerDeadzone,
        activeConfig.stickDeadzone,
      );
      if (elapsedMs > 0 && Math.abs(primaryScroll) >= activeConfig.stickDeadzone) {
        options.scrollPopup(
          (primaryScroll * activeConfig.scrollPixelsPerSecond * elapsedMs) / 1000,
        );
      }

      handleJumpAxis(
        resolveAxisBindingValue(
          axisGamepad,
          activeConfig.bindings.rightStickVertical,
          activeConfig.triggerDeadzone,
          activationThreshold,
        ),
        now,
        activeConfig,
      );
    } else {
      resetHeldAction(jumpHold);
      for (const actionKey of POPUP_ONLY_ACTIONS) {
        trackActionWithoutFiring(actionKey, activeGamepad, activeConfig);
      }
    }

    handleSelectionAxis(
      resolveAxisBindingValue(
        axisGamepad,
        activeConfig.bindings.leftStickHorizontal,
        activeConfig.triggerDeadzone,
        activationThreshold,
      ),
      now,
      activeConfig,
    );
  }

  return {
    poll,
    getActiveGamepadId: (): string | null => activeGamepadId,
  };
}
