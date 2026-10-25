import {
  compileControllerReviewBindings,
  formatControllerReviewBinding,
  getControllerReviewAction,
  CONTROLLER_REVIEW_ACTIONS,
  type CompiledControllerReviewBinding,
  type ControllerReviewInputContext,
  type ControllerReviewInputSource,
  type ControllerReviewStep,
} from '../../shared/controller-review';
import type { ControllerReviewAction, ResolvedControllerConfig } from '../../types';
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

/** Every button on the media timing review screen, named by what it does. */
export type MediaTimingReviewControllerAction = ControllerReviewAction;

/** The review shows either the timing editor or the "Stop reviewing?" step. */
export type MediaTimingReviewStep = ControllerReviewStep;

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

/** Hint text per action for the review buttons, e.g. `A` or `L1 + ▶`; null when unbound. */
export function formatMediaTimingReviewControllerHints(
  config: ResolvedControllerConfig,
  mapping: string | null | undefined,
): Record<MediaTimingReviewControllerAction, string | null> {
  const context = getControllerReviewInputContext(config, mapping);
  const hints = {} as Record<MediaTimingReviewControllerAction, string | null>;
  for (const action of CONTROLLER_REVIEW_ACTIONS) {
    const binding = config.reviewBindings[action.id];
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

type PressChannel = {
  action: MediaTimingReviewControllerAction | null;
  repeat: boolean;
  /** Hold buttons that were down when this press started. */
  modifiers: string;
  /** Set when the hold buttons changed mid-press; the press must be released first. */
  latched: boolean;
  hold: HoldState;
};

type ReviewFrame = {
  compiled: CompiledControllerReviewBinding[];
  pressed: Set<string>;
  holdKeys: Set<string>;
};

/**
 * Picks the action for a fresh press: with one hold button down, the binding that uses it;
 * with none, the plain binding. Two hold buttons at once is ambiguous and matches nothing.
 */
function matchAction(
  compiled: CompiledControllerReviewBinding[],
  pressKey: string,
  modifiers: string[],
): MediaTimingReviewControllerAction | null {
  if (modifiers.length > 1) return null;
  const holdKey = modifiers[0] ?? null;
  return (
    compiled.find((entry) => entry.pressKey === pressKey && entry.holdKey === holdKey)?.action ??
    null
  );
}

export function createMediaTimingReviewControllerLayout(options: {
  getStep: () => MediaTimingReviewStep;
  dispatch: (action: MediaTimingReviewControllerAction) => void;
}): ControllerScreenLayout {
  const channels = new Map<string, PressChannel>();

  function readFrame(gamepad: GamepadLike, config: ResolvedControllerConfig): ReviewFrame {
    const context = getControllerReviewInputContext(config, gamepad.mapping);
    const step = options.getStep();
    const compiled = compileControllerReviewBindings(config.reviewBindings, step, context);
    // Track inputs from both steps so a press held across a step change is not new.
    const sources = new Map<string, ControllerReviewInputSource>();
    for (const anyStep of ['editor', 'cancel'] as const) {
      const entries =
        anyStep === step
          ? compiled
          : compileControllerReviewBindings(config.reviewBindings, anyStep, context);
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

  function heldModifiers(frame: ReviewFrame, exceptKey: string): string[] {
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
        const next: PressChannel = {
          action,
          repeat: action ? getControllerReviewAction(action).repeat : false,
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
