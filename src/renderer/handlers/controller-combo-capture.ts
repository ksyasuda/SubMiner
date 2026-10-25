import type { ControllerReviewInput, ResolvedControllerReviewBinding } from '../../types';

type ControllerButtonState = {
  value: number;
  pressed?: boolean;
  touched?: boolean;
};

type ControllerSnapshot = {
  axes: readonly number[];
  buttons: readonly ControllerButtonState[];
};

export type ControllerComboCaptureState =
  | { status: 'waiting' }
  | { status: 'holding'; hold: ControllerReviewInput }
  | { status: 'done'; binding: ResolvedControllerReviewBinding };

type RawInput = Extract<ControllerReviewInput, { kind: 'button' | 'axis' }>;

function inputKey(input: RawInput): string {
  return input.kind === 'button'
    ? `button:${input.buttonIndex}`
    : `axis:${input.axisIndex}:${input.direction}`;
}

/**
 * Learn for review bindings. Waits for a first input; letting go of it records it alone,
 * while pressing a second input during the hold records "hold first + press second".
 * Buttons already down when learning starts are ignored until released. Axis directions
 * active at the start stay ignored for the whole capture: some pads rest a trigger axis at
 * -1, and brushing it must not record that resting position as a press.
 */
export function createControllerComboCapture(options: {
  triggerDeadzone: number;
  stickDeadzone: number;
}) {
  const blocked = new Set<string>();
  const restingAxes = new Set<string>();
  let first: RawInput | null = null;

  function activeInputs(snapshot: ControllerSnapshot): RawInput[] {
    const inputs: RawInput[] = [];
    snapshot.buttons.forEach((button, buttonIndex) => {
      if (button && (Boolean(button.pressed) || button.value >= options.triggerDeadzone)) {
        inputs.push({ kind: 'button', buttonIndex });
      }
    });
    const threshold = Math.max(options.stickDeadzone, 0.55);
    snapshot.axes.forEach((value, axisIndex) => {
      if (typeof value !== 'number' || !Number.isFinite(value) || Math.abs(value) < threshold) {
        return;
      }
      inputs.push({ kind: 'axis', axisIndex, direction: value > 0 ? 'positive' : 'negative' });
    });
    return inputs;
  }

  function arm(snapshot: ControllerSnapshot): void {
    first = null;
    blocked.clear();
    restingAxes.clear();
    for (const input of activeInputs(snapshot)) {
      (input.kind === 'axis' ? restingAxes : blocked).add(inputKey(input));
    }
  }

  function poll(snapshot: ControllerSnapshot): ControllerComboCaptureState {
    const active = activeInputs(snapshot).filter((input) => !restingAxes.has(inputKey(input)));
    const activeKeys = new Set(active.map(inputKey));
    for (const key of [...blocked]) {
      if (!activeKeys.has(key)) blocked.delete(key);
    }
    const fresh = active.filter((input) => !blocked.has(inputKey(input)));

    if (first === null) {
      const next = fresh[0];
      if (!next) return { status: 'waiting' };
      first = next;
      return { status: 'holding', hold: first };
    }

    const firstKey = inputKey(first);
    if (!activeKeys.has(firstKey)) {
      const press = first;
      first = null;
      return { status: 'done', binding: { kind: 'combo', press, hold: null } };
    }
    const second = fresh.find((input) => inputKey(input) !== firstKey);
    if (second) {
      const hold = first;
      first = null;
      return { status: 'done', binding: { kind: 'combo', press: second, hold } };
    }
    return { status: 'holding', hold: first };
  }

  return { arm, poll };
}
