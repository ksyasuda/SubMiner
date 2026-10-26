import {
  compileControllerScreenBindings,
  formatControllerReviewBinding,
  type CompiledControllerScreenBinding,
  type ControllerReviewInputContext,
  type ControllerReviewInputSource,
  type ControllerScreenBindings,
  type ControllerScreenDefinition,
} from '../../shared/controller-review';
import type { ResolvedControllerConfig } from '../../types';
import {
  createHoldState,
  resolveDiscreteBindingPressed,
  resolveDpadHorizontalValue,
  resolveDpadVerticalValue,
  shouldFireHeldAction,
  type ControllerScreenLayout,
  type GamepadLike,
  type HoldState,
} from './gamepad-controller.js';

export function getControllerReviewInputContext(
  config: ResolvedControllerConfig,
  mapping: string | null | undefined,
): ControllerReviewInputContext {
  return {
    mapping,
    faceButtonLabels: config.faceButtonLabels,
    buttonIndices: config.buttonIndices,
  };
}

/** Hint text per action of a screen, e.g. `A` or `L1 + ▶`; null when unbound. */
export function formatControllerScreenHints<A extends string>(
  screen: ControllerScreenDefinition<A>,
  bindings: Partial<ControllerScreenBindings<A>>,
  config: ResolvedControllerConfig,
  mapping: string | null | undefined,
): Record<A, string | null> {
  const context = getControllerReviewInputContext(config, mapping);
  const hints = {} as Record<A, string | null>;
  for (const action of screen.actions) {
    const binding = bindings[action.id];
    hints[action.id] =
      binding && binding.kind === 'combo' ? formatControllerReviewBinding(binding, context) : null;
  }
  return hints;
}

function isSourcePressed(
  source: ControllerReviewInputSource,
  gamepad: GamepadLike,
  config: ResolvedControllerConfig,
): boolean {
  if (source.kind === 'dpad') {
    if (source.direction === 'left' || source.direction === 'right') {
      const value = resolveDpadHorizontalValue(gamepad, config.triggerDeadzone);
      return source.direction === 'left' ? value < 0 : value > 0;
    }
    const value = resolveDpadVerticalValue(gamepad, config.triggerDeadzone);
    return source.direction === 'up' ? value < 0 : value > 0;
  }
  return resolveDiscreteBindingPressed(gamepad, source, config);
}

type PressChannel<A extends string> = {
  action: A | null;
  repeat: boolean;
  /** Hold buttons that were down when this press started. */
  modifiers: string;
  /** Set when the hold buttons changed mid-press; the press must be released first. */
  latched: boolean;
  hold: HoldState;
};

type ScreenFrame<A extends string> = {
  compiled: CompiledControllerScreenBinding<A>[];
  pressed: Set<string>;
  holdKeys: Set<string>;
};

/**
 * Picks the action for a fresh press: with one hold button down, the binding that uses it;
 * with none, the plain binding. Two hold buttons at once is ambiguous and matches nothing.
 */
function matchAction<A extends string>(
  compiled: CompiledControllerScreenBinding<A>[],
  pressKey: string,
  modifiers: string[],
): A | null {
  if (modifiers.length > 1) return null;
  const holdKey = modifiers[0] ?? null;
  return (
    compiled.find((entry) => entry.pressKey === pressKey && entry.holdKey === holdKey)?.action ??
    null
  );
}

/**
 * Controller layout for a dialog described by a screen definition: per step, press-or-combo
 * bindings with optional repeat. Actions are dispatched; the dialog decides what they do.
 */
export function createControllerComboScreenLayout<A extends string, S extends string>(options: {
  screen: ControllerScreenDefinition<A, S>;
  getBindings: (config: ResolvedControllerConfig) => Partial<ControllerScreenBindings<A>>;
  getStep: () => S;
  dispatch: (action: A) => void;
}): ControllerScreenLayout {
  const { screen } = options;
  const channels = new Map<string, PressChannel<A>>();
  const repeatByAction = new Map(screen.actions.map((action) => [action.id, action.repeat]));

  function readFrame(gamepad: GamepadLike, config: ResolvedControllerConfig): ScreenFrame<A> {
    const context = getControllerReviewInputContext(config, gamepad.mapping);
    const step = options.getStep();
    const bindings = options.getBindings(config);
    const compiled = compileControllerScreenBindings(screen, bindings, step, context);
    // Track inputs from every step so a press held across a step change is not new.
    const sources = new Map<string, ControllerReviewInputSource>();
    for (const { id: anyStep } of screen.steps) {
      const entries =
        anyStep === step
          ? compiled
          : compileControllerScreenBindings(screen, bindings, anyStep, context);
      for (const entry of entries) {
        sources.set(entry.pressKey, entry.press);
        if (entry.hold && entry.holdKey) sources.set(entry.holdKey, entry.hold);
      }
    }
    const pressed = new Set<string>();
    for (const [key, source] of sources) {
      if (isSourcePressed(source, gamepad, config)) pressed.add(key);
    }
    const holdKeys = new Set<string>();
    for (const entry of compiled) {
      if (entry.holdKey) holdKeys.add(entry.holdKey);
    }
    return { compiled, pressed, holdKeys };
  }

  function heldModifiers(frame: ScreenFrame<A>, exceptKey: string): string[] {
    return [...frame.holdKeys].filter((key) => key !== exceptKey && frame.pressed.has(key)).sort();
  }

  function poll(gamepad: GamepadLike, config: ResolvedControllerConfig, now: number): void {
    const frame = readFrame(gamepad, config);

    for (const key of [...channels.keys()]) {
      if (!frame.pressed.has(key)) channels.delete(key);
    }

    for (const key of frame.pressed) {
      const modifiers = heldModifiers(frame, key);
      const signature = modifiers.join('|');
      const channel = channels.get(key);

      if (!channel) {
        const action = matchAction(frame.compiled, key, modifiers);
        const next: PressChannel<A> = {
          action,
          repeat: action ? (repeatByAction.get(action) ?? false) : false,
          modifiers: signature,
          latched: false,
          hold: createHoldState(),
        };
        channels.set(key, next);
        if (action) {
          shouldFireHeldAction(next.hold, now, config.repeatDelayMs, config.repeatIntervalMs);
          options.dispatch(action);
        }
        continue;
      }

      if (channel.latched || channel.action === null) continue;
      if (channel.modifiers !== signature) {
        // Letting go of L1 a moment before the D-pad must not fire the plain D-pad action.
        channel.latched = true;
        continue;
      }
      if (
        channel.repeat &&
        shouldFireHeldAction(channel.hold, now, config.repeatDelayMs, config.repeatIntervalMs)
      ) {
        options.dispatch(channel.action);
      }
    }
  }

  function latchHeld(gamepad: GamepadLike, config: ResolvedControllerConfig): void {
    channels.clear();
    for (const key of readFrame(gamepad, config).pressed) {
      channels.set(key, {
        action: null,
        repeat: false,
        modifiers: '',
        latched: true,
        hold: createHoldState(),
      });
    }
  }

  return {
    activate(gamepad, config, now, { latchHeld: shouldLatch }) {
      if (shouldLatch) {
        latchHeld(gamepad, config);
        return;
      }
      channels.clear();
      poll(gamepad, config, now);
    },
    poll,
  };
}
