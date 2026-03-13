import type {
  ControllerDpadFallback,
  ResolvedControllerAxisBinding,
  ResolvedControllerDiscreteBinding,
} from '../../types';

type ControllerButtonState = {
  value: number;
  pressed?: boolean;
  touched?: boolean;
};

type ControllerBindingCaptureSnapshot = {
  axes: readonly number[];
  buttons: readonly ControllerButtonState[];
};

type ControllerBindingCaptureTarget =
  | {
      actionId: string;
      bindingType: 'discrete';
    }
  | {
      actionId: string;
      bindingType: 'axis';
      dpadFallback: ControllerDpadFallback;
    };

type ControllerBindingCaptureResult =
  | {
      actionId: string;
      bindingType: 'discrete';
      binding: ResolvedControllerDiscreteBinding;
    }
  | {
      actionId: string;
      bindingType: 'axis';
      binding: ResolvedControllerAxisBinding;
    };

function isActiveButton(button: ControllerButtonState | undefined, triggerDeadzone: number): boolean {
  if (!button) return false;
  return Boolean(button.pressed) || button.value >= triggerDeadzone;
}

function getAxisDirection(
  value: number | undefined,
  activationThreshold: number,
): 'negative' | 'positive' | null {
  if (typeof value !== 'number' || !Number.isFinite(value)) return null;
  if (Math.abs(value) < activationThreshold) return null;
  return value > 0 ? 'positive' : 'negative';
}

export function createControllerBindingCapture(options: {
  triggerDeadzone: number;
  stickDeadzone: number;
}) {
  let target: ControllerBindingCaptureTarget | null = null;
  const blockedButtons = new Set<number>();
  const blockedAxisDirections = new Set<string>();

  function resetBlockedState(snapshot: ControllerBindingCaptureSnapshot): void {
    blockedButtons.clear();
    blockedAxisDirections.clear();

    snapshot.buttons.forEach((button, index) => {
      if (isActiveButton(button, options.triggerDeadzone)) {
        blockedButtons.add(index);
      }
    });

    const activationThreshold = Math.max(options.stickDeadzone, 0.55);
    snapshot.axes.forEach((value, index) => {
      const direction = getAxisDirection(value, activationThreshold);
      if (direction) {
        blockedAxisDirections.add(`${index}:${direction}`);
      }
    });
  }

  function arm(nextTarget: ControllerBindingCaptureTarget, snapshot: ControllerBindingCaptureSnapshot): void {
    target = nextTarget;
    resetBlockedState(snapshot);
  }

  function cancel(): void {
    target = null;
    blockedButtons.clear();
    blockedAxisDirections.clear();
  }

  function poll(snapshot: ControllerBindingCaptureSnapshot): ControllerBindingCaptureResult | null {
    if (!target) return null;

    snapshot.buttons.forEach((button, index) => {
      if (!isActiveButton(button, options.triggerDeadzone)) {
        blockedButtons.delete(index);
      }
    });

    const activationThreshold = Math.max(options.stickDeadzone, 0.55);
    snapshot.axes.forEach((value, index) => {
      const negativeKey = `${index}:negative`;
      const positiveKey = `${index}:positive`;
      if (getAxisDirection(value, activationThreshold) === null) {
        blockedAxisDirections.delete(negativeKey);
        blockedAxisDirections.delete(positiveKey);
      }
    });

    for (let index = 0; index < snapshot.buttons.length; index += 1) {
      if (!isActiveButton(snapshot.buttons[index], options.triggerDeadzone)) continue;
      if (blockedButtons.has(index)) continue;

      const result: ControllerBindingCaptureResult =
        target.bindingType === 'discrete'
          ? {
              actionId: target.actionId,
              bindingType: 'discrete',
              binding: { kind: 'button', buttonIndex: index },
            }
          : {
              actionId: target.actionId,
              bindingType: 'axis',
              binding: {
                kind: 'axis',
                axisIndex: index,
                dpadFallback: target.dpadFallback,
              },
            };
      cancel();
      return result;
    }

    for (let index = 0; index < snapshot.axes.length; index += 1) {
      const direction = getAxisDirection(snapshot.axes[index], activationThreshold);
      if (!direction) continue;
      const directionKey = `${index}:${direction}`;
      if (blockedAxisDirections.has(directionKey)) continue;

      const result: ControllerBindingCaptureResult =
        target.bindingType === 'discrete'
          ? {
              actionId: target.actionId,
              bindingType: 'discrete',
              binding: { kind: 'axis', axisIndex: index, direction },
            }
          : {
              actionId: target.actionId,
              bindingType: 'axis',
              binding: {
                kind: 'axis',
                axisIndex: index,
                dpadFallback: target.dpadFallback,
              },
            };
      cancel();
      return result;
    }

    return null;
  }

  return {
    arm,
    cancel,
    isArmed: (): boolean => target !== null,
    getTargetActionId: (): string | null => target?.actionId ?? null,
    poll,
  };
}
