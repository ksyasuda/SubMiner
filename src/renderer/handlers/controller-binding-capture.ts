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
    }
  | {
      actionId: string;
      bindingType: 'dpad';
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
    }
  | {
      actionId: string;
      bindingType: 'dpad';
      dpadDirection: ControllerDpadFallback;
    };

function isActiveButton(
  button: ControllerButtonState | undefined,
  triggerDeadzone: number,
): boolean {
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

const DPAD_BUTTON_INDICES = [12, 13, 14, 15] as const;

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

  function arm(
    nextTarget: ControllerBindingCaptureTarget,
    snapshot: ControllerBindingCaptureSnapshot,
  ): void {
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

    // D-pad capture: only respond to d-pad buttons (12-15)
    if (target.bindingType === 'dpad') {
      for (const index of DPAD_BUTTON_INDICES) {
        if (!isActiveButton(snapshot.buttons[index], options.triggerDeadzone)) continue;
        if (blockedButtons.has(index)) continue;

        const dpadDirection: ControllerDpadFallback =
          index === 12 || index === 13 ? 'vertical' : 'horizontal';
        cancel();
        return {
          actionId: target.actionId,
          bindingType: 'dpad' as const,
          dpadDirection,
        };
      }
      return null;
    }

    // After dpad early-return, only 'discrete' | 'axis' remain
    const narrowedTarget: Extract<
      ControllerBindingCaptureTarget,
      { bindingType: 'discrete' | 'axis' }
    > = target;

    for (let index = 0; index < snapshot.buttons.length; index += 1) {
      if (!isActiveButton(snapshot.buttons[index], options.triggerDeadzone)) continue;
      if (blockedButtons.has(index)) continue;
      if (narrowedTarget.bindingType === 'axis') continue;

      const result: ControllerBindingCaptureResult = {
        actionId: narrowedTarget.actionId,
        bindingType: 'discrete',
        binding: { kind: 'button', buttonIndex: index },
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
        narrowedTarget.bindingType === 'discrete'
          ? {
              actionId: narrowedTarget.actionId,
              bindingType: 'discrete',
              binding: { kind: 'axis', axisIndex: index, direction },
            }
          : {
              actionId: narrowedTarget.actionId,
              bindingType: 'axis',
              binding: {
                kind: 'axis',
                axisIndex: index,
                dpadFallback: narrowedTarget.dpadFallback,
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
