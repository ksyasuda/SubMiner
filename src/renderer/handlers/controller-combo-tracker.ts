import { controllerReviewInputSourceKey as inputKey } from '../../shared/controller-review';
import type { ControllerRawInput, ResolvedControllerDiscreteBinding } from '../../types';
import type { GamepadLike } from './gamepad-controller.js';

type BoundBinding = Exclude<ResolvedControllerDiscreteBinding, { kind: 'none' }>;

const isBound = (binding: ResolvedControllerDiscreteBinding): binding is BoundBinding =>
  binding.kind !== 'none';

/**
 * Per-poll view of hold + press combos in the main controller bindings. Dialog layouts have
 * their own matching in `controller-screen-layout.ts`; the rules for choosing a binding match.
 *
 * A press belongs to the binding that matches the hold buttons down when it starts, so
 * pressing the same input again during one hold fires the combo again, and letting go of the
 * hold first does not turn the press into its plain binding.
 */
export function createControllerComboTracker() {
  let pressedLastPoll: ReadonlySet<string> = new Set();
  let holdKeys: ReadonlySet<string> = new Set();
  let heldHoldKeys: string[] = [];
  let usedHolds: ReadonlySet<string> = new Set();
  /** Holds used during their press, kept for the poll that sees them released. */
  let cancelledTaps: ReadonlySet<string> = new Set();
  const claimedInputs = new Map<string, ControllerRawInput>();

  function update(
    bindings: readonly ResolvedControllerDiscreteBinding[],
    isPressed: (input: ControllerRawInput) => boolean,
  ): void {
    const bound = bindings.filter(isBound);
    const pressed = new Set<string>();
    for (const binding of bound) {
      for (const input of binding.hold ? [binding, binding.hold] : [binding]) {
        if (isPressed(input)) pressed.add(inputKey(input));
      }
    }

    holdKeys = new Set(bound.flatMap((binding) => (binding.hold ? [inputKey(binding.hold)] : [])));
    heldHoldKeys = [...holdKeys].filter((key) => pressed.has(key));

    const freshPresses = bound
      .map(inputKey)
      .filter((key) => pressed.has(key) && !pressedLastPoll.has(key));
    const nextUsedHolds = new Set(
      heldHoldKeys.filter(
        (holdKey) => usedHolds.has(holdKey) || freshPresses.some((key) => key !== holdKey),
      ),
    );
    cancelledTaps = new Set([...usedHolds, ...nextUsedHolds]);
    usedHolds = nextUsedHolds;
    pressedLastPoll = pressed;

    for (const key of claimedInputs.keys()) {
      if (!pressed.has(key)) claimedInputs.delete(key);
    }
    for (const binding of bound) {
      if (!binding.hold) continue;
      const pressKey = inputKey(binding);
      const holdKey = inputKey(binding.hold);
      if (pressKey !== holdKey && pressed.has(pressKey) && pressed.has(holdKey)) {
        claimedInputs.set(pressKey, binding);
      }
    }
  }

  return {
    update,

    reset(): void {
      pressedLastPoll = new Set();
      holdKeys = new Set();
      heldHoldKeys = [];
      usedHolds = new Set();
      cancelledTaps = new Set();
      claimedInputs.clear();
    },

    /**
     * Whether a press of this binding counts right now: with one hold button down, only the
     * binding that uses it; with none, only a plain binding. Two holds down match nothing.
     */
    matches(binding: ResolvedControllerDiscreteBinding): boolean {
      if (!isBound(binding)) return false;
      const pressKey = inputKey(binding);
      const others = heldHoldKeys.filter((key) => key !== pressKey);
      if (others.length > 1) return false;
      return others[0] === (binding.hold ? inputKey(binding.hold) : undefined);
    },

    /** A plain binding on a hold button waits for release, when a tap can be told from a hold. */
    firesOnRelease(binding: ResolvedControllerDiscreteBinding): boolean {
      return isBound(binding) && !binding.hold && holdKeys.has(inputKey(binding));
    },

    /** Another bound input was pressed during this hold, so its release is not a tap. */
    isTapCancelled(binding: ResolvedControllerDiscreteBinding): boolean {
      return isBound(binding) && cancelledTaps.has(inputKey(binding));
    },

    /**
     * The pad with inputs claimed by a combo reading as released, for the stick actions and
     * their D-pad fallbacks. A claim lasts until the input is released, so letting go of the
     * hold first does not start moving the word selection.
     */
    withoutClaimedInputs(gamepad: GamepadLike): GamepadLike {
      if (claimedInputs.size === 0) return gamepad;
      const axes = [...gamepad.axes];
      const buttons = [...gamepad.buttons];
      for (const input of claimedInputs.values()) {
        if (input.kind === 'button') {
          buttons[input.buttonIndex] = { value: 0, pressed: false, touched: false };
        } else if ((axes[input.axisIndex] ?? 0) > 0 === (input.direction === 'positive')) {
          axes[input.axisIndex] = 0;
        }
      }
      // Native Gamepad fields are prototype getters, so object spread would drop them.
      return {
        id: gamepad.id,
        index: gamepad.index,
        connected: gamepad.connected,
        mapping: gamepad.mapping,
        timestamp: gamepad.timestamp,
        axes,
        buttons,
      };
    },
  };
}
