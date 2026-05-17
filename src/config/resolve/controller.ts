import type {
  ControllerAxisBinding,
  ControllerAxisDirection,
  ControllerButtonBinding,
  ControllerButtonIndicesConfig,
  ControllerDpadFallback,
  ResolvedControllerAxisBinding,
  ResolvedControllerBindingsConfig,
  ResolvedControllerDiscreteBinding,
} from '../../types/runtime';
import { ResolveContext } from './context';
import { asBoolean, asNumber, asString, isObject } from './shared';

const CONTROLLER_BUTTON_BINDINGS = [
  'none',
  'select',
  'buttonSouth',
  'buttonEast',
  'buttonNorth',
  'buttonWest',
  'leftShoulder',
  'rightShoulder',
  'leftStickPress',
  'rightStickPress',
  'leftTrigger',
  'rightTrigger',
] as const;

const CONTROLLER_AXIS_BINDINGS = [
  'leftStickX',
  'leftStickY',
  'rightStickX',
  'rightStickY',
] as const;

const CONTROLLER_AXIS_INDEX_BY_BINDING: Record<ControllerAxisBinding, number> = {
  leftStickX: 0,
  leftStickY: 1,
  rightStickX: 3,
  rightStickY: 4,
};

const CONTROLLER_BUTTON_INDEX_KEY_BY_BINDING: Record<
  Exclude<ControllerButtonBinding, 'none'>,
  keyof Required<ControllerButtonIndicesConfig>
> = {
  select: 'select',
  buttonSouth: 'buttonSouth',
  buttonEast: 'buttonEast',
  buttonNorth: 'buttonNorth',
  buttonWest: 'buttonWest',
  leftShoulder: 'leftShoulder',
  rightShoulder: 'rightShoulder',
  leftStickPress: 'leftStickPress',
  rightStickPress: 'rightStickPress',
  leftTrigger: 'leftTrigger',
  rightTrigger: 'rightTrigger',
};

const CONTROLLER_AXIS_FALLBACK_BY_SLOT = {
  leftStickHorizontal: 'horizontal',
  leftStickVertical: 'vertical',
  rightStickHorizontal: 'none',
  rightStickVertical: 'none',
} as const satisfies Record<string, ControllerDpadFallback>;

const CONTROLLER_BUTTON_INDEX_KEYS = [
  'select',
  'buttonSouth',
  'buttonEast',
  'buttonNorth',
  'buttonWest',
  'leftShoulder',
  'rightShoulder',
  'leftStickPress',
  'rightStickPress',
  'leftTrigger',
  'rightTrigger',
] as const;

const CONTROLLER_DISCRETE_BINDING_KEYS = [
  'toggleLookup',
  'closeLookup',
  'toggleKeyboardOnlyMode',
  'mineCard',
  'quitMpv',
  'previousAudio',
  'nextAudio',
  'playCurrentAudio',
  'toggleMpvPause',
] as const;

const CONTROLLER_AXIS_BINDING_KEYS = [
  'leftStickHorizontal',
  'leftStickVertical',
  'rightStickHorizontal',
  'rightStickVertical',
] as const;

type ControllerBindingsTarget = Required<ResolvedControllerBindingsConfig>;
type ControllerButtonIndicesTarget = Required<ControllerButtonIndicesConfig>;

function isControllerAxisDirection(value: unknown): value is ControllerAxisDirection {
  return value === 'negative' || value === 'positive';
}

function isControllerDpadFallback(value: unknown): value is ControllerDpadFallback {
  return value === 'none' || value === 'horizontal' || value === 'vertical';
}

function resolveLegacyDiscreteBinding(
  value: ControllerButtonBinding,
  buttonIndices: Required<ControllerButtonIndicesConfig>,
): ResolvedControllerDiscreteBinding {
  if (value === 'none') {
    return { kind: 'none' };
  }
  return {
    kind: 'button',
    buttonIndex: buttonIndices[CONTROLLER_BUTTON_INDEX_KEY_BY_BINDING[value]],
  };
}

function resolveLegacyAxisBinding(
  value: ControllerAxisBinding,
  slot: keyof typeof CONTROLLER_AXIS_FALLBACK_BY_SLOT,
): ResolvedControllerAxisBinding {
  return {
    kind: 'axis',
    axisIndex: CONTROLLER_AXIS_INDEX_BY_BINDING[value],
    dpadFallback: CONTROLLER_AXIS_FALLBACK_BY_SLOT[slot],
  };
}

function parseDiscreteBindingObject(value: unknown): ResolvedControllerDiscreteBinding | null {
  if (!isObject(value) || typeof value.kind !== 'string') return null;
  if (value.kind === 'none') {
    return { kind: 'none' };
  }
  if (value.kind === 'button') {
    return typeof value.buttonIndex === 'number' &&
      Number.isInteger(value.buttonIndex) &&
      value.buttonIndex >= 0
      ? { kind: 'button', buttonIndex: value.buttonIndex }
      : null;
  }
  if (value.kind === 'axis') {
    return typeof value.axisIndex === 'number' &&
      Number.isInteger(value.axisIndex) &&
      value.axisIndex >= 0 &&
      isControllerAxisDirection(value.direction)
      ? { kind: 'axis', axisIndex: value.axisIndex, direction: value.direction }
      : null;
  }
  return null;
}

function parseAxisBindingObject(
  value: unknown,
  slot: keyof typeof CONTROLLER_AXIS_FALLBACK_BY_SLOT,
): ResolvedControllerAxisBinding | null {
  if (isObject(value) && value.kind === 'none') {
    return { kind: 'none' };
  }
  if (!isObject(value) || value.kind !== 'axis') return null;
  if (
    typeof value.axisIndex !== 'number' ||
    !Number.isInteger(value.axisIndex) ||
    value.axisIndex < 0
  ) {
    return null;
  }
  if (value.dpadFallback !== undefined && !isControllerDpadFallback(value.dpadFallback)) {
    return null;
  }
  return {
    kind: 'axis',
    axisIndex: value.axisIndex,
    dpadFallback: value.dpadFallback ?? CONTROLLER_AXIS_FALLBACK_BY_SLOT[slot],
  };
}

function applyControllerButtonIndices(
  source: unknown,
  target: ControllerButtonIndicesTarget,
  pathPrefix: string,
  warn: ResolveContext['warn'],
): void {
  if (!isObject(source)) return;

  for (const key of CONTROLLER_BUTTON_INDEX_KEYS) {
    const value = asNumber(source[key]);
    if (value !== undefined && value >= 0 && Number.isInteger(value)) {
      target[key] = value;
    } else if (source[key] !== undefined) {
      warn(`${pathPrefix}.${key}`, source[key], target[key], 'Expected non-negative integer.');
    }
  }
}

function applyControllerBindings(
  source: unknown,
  target: ControllerBindingsTarget,
  buttonIndices: ControllerButtonIndicesTarget,
  pathPrefix: string,
  warn: ResolveContext['warn'],
): void {
  if (!isObject(source)) return;

  for (const key of CONTROLLER_DISCRETE_BINDING_KEYS) {
    const bindingValue = source[key];
    const legacyValue = asString(bindingValue);
    if (
      legacyValue !== undefined &&
      CONTROLLER_BUTTON_BINDINGS.includes(
        legacyValue as (typeof CONTROLLER_BUTTON_BINDINGS)[number],
      )
    ) {
      target[key] = resolveLegacyDiscreteBinding(
        legacyValue as ControllerButtonBinding,
        buttonIndices,
      );
      continue;
    }
    const parsedObject = parseDiscreteBindingObject(bindingValue);
    if (parsedObject) {
      target[key] = parsedObject;
    } else if (bindingValue !== undefined) {
      warn(
        `${pathPrefix}.${key}`,
        bindingValue,
        target[key],
        "Expected legacy controller button name or binding object with kind 'none', 'button', or 'axis'.",
      );
    }
  }

  for (const key of CONTROLLER_AXIS_BINDING_KEYS) {
    const bindingValue = source[key];
    const legacyValue = asString(bindingValue);
    if (
      legacyValue !== undefined &&
      CONTROLLER_AXIS_BINDINGS.includes(legacyValue as (typeof CONTROLLER_AXIS_BINDINGS)[number])
    ) {
      target[key] = resolveLegacyAxisBinding(legacyValue as ControllerAxisBinding, key);
      continue;
    }
    if (legacyValue === 'none') {
      target[key] = { kind: 'none' };
      continue;
    }
    const parsedObject = parseAxisBindingObject(bindingValue, key);
    if (parsedObject) {
      target[key] = parsedObject;
    } else if (bindingValue !== undefined) {
      warn(
        `${pathPrefix}.${key}`,
        bindingValue,
        target[key],
        "Expected legacy controller axis name ('none' allowed) or binding object with kind 'axis'.",
      );
    }
  }
}

export function applyControllerConfig(context: ResolveContext): void {
  const { src, resolved, warn } = context;
  if (!isObject(src.controller)) return;

  const enabled = asBoolean(src.controller.enabled);
  if (enabled !== undefined) {
    resolved.controller.enabled = enabled;
  } else if (src.controller.enabled !== undefined) {
    warn(
      'controller.enabled',
      src.controller.enabled,
      resolved.controller.enabled,
      'Expected boolean.',
    );
  }

  const preferredGamepadId = asString(src.controller.preferredGamepadId);
  if (preferredGamepadId !== undefined) {
    resolved.controller.preferredGamepadId = preferredGamepadId;
  }

  const preferredGamepadLabel = asString(src.controller.preferredGamepadLabel);
  if (preferredGamepadLabel !== undefined) {
    resolved.controller.preferredGamepadLabel = preferredGamepadLabel;
  }

  const smoothScroll = asBoolean(src.controller.smoothScroll);
  if (smoothScroll !== undefined) {
    resolved.controller.smoothScroll = smoothScroll;
  } else if (src.controller.smoothScroll !== undefined) {
    warn(
      'controller.smoothScroll',
      src.controller.smoothScroll,
      resolved.controller.smoothScroll,
      'Expected boolean.',
    );
  }

  const triggerInputMode = asString(src.controller.triggerInputMode);
  if (
    triggerInputMode === 'auto' ||
    triggerInputMode === 'digital' ||
    triggerInputMode === 'analog'
  ) {
    resolved.controller.triggerInputMode = triggerInputMode;
  } else if (src.controller.triggerInputMode !== undefined) {
    warn(
      'controller.triggerInputMode',
      src.controller.triggerInputMode,
      resolved.controller.triggerInputMode,
      "Expected 'auto', 'digital', or 'analog'.",
    );
  }

  const boundedNumberKeys = [
    'scrollPixelsPerSecond',
    'horizontalJumpPixels',
    'repeatDelayMs',
    'repeatIntervalMs',
  ] as const;
  for (const key of boundedNumberKeys) {
    const value = asNumber(src.controller[key]);
    if (value !== undefined && Math.floor(value) > 0) {
      resolved.controller[key] = Math.floor(value) as (typeof resolved.controller)[typeof key];
    } else if (src.controller[key] !== undefined) {
      warn(
        `controller.${key}`,
        src.controller[key],
        resolved.controller[key],
        'Expected positive number.',
      );
    }
  }

  const deadzoneKeys = ['stickDeadzone', 'triggerDeadzone'] as const;
  for (const key of deadzoneKeys) {
    const value = asNumber(src.controller[key]);
    if (value !== undefined && value >= 0 && value <= 1) {
      resolved.controller[key] = value as (typeof resolved.controller)[typeof key];
    } else if (src.controller[key] !== undefined) {
      warn(
        `controller.${key}`,
        src.controller[key],
        resolved.controller[key],
        'Expected number between 0 and 1.',
      );
    }
  }

  applyControllerButtonIndices(
    src.controller.buttonIndices,
    resolved.controller.buttonIndices,
    'controller.buttonIndices',
    warn,
  );
  applyControllerBindings(
    src.controller.bindings,
    resolved.controller.bindings,
    resolved.controller.buttonIndices,
    'controller.bindings',
    warn,
  );

  if (isObject(src.controller.profiles)) {
    for (const [profileId, rawProfile] of Object.entries(src.controller.profiles)) {
      if (!isObject(rawProfile)) {
        warn(
          `controller.profiles.${profileId}`,
          rawProfile,
          undefined,
          'Expected controller profile object.',
        );
        continue;
      }

      const label = asString(rawProfile.label);
      if (rawProfile.label !== undefined && label === undefined) {
        warn(
          `controller.profiles.${profileId}.label`,
          rawProfile.label,
          profileId,
          'Expected string.',
        );
      }

      const profile = {
        label: label ?? profileId,
        buttonIndices: structuredClone(resolved.controller.buttonIndices),
        bindings: structuredClone(resolved.controller.bindings),
      };
      applyControllerButtonIndices(
        rawProfile.buttonIndices,
        profile.buttonIndices,
        `controller.profiles.${profileId}.buttonIndices`,
        warn,
      );
      applyControllerBindings(
        rawProfile.bindings,
        profile.bindings,
        profile.buttonIndices,
        `controller.profiles.${profileId}.bindings`,
        warn,
      );
      resolved.controller.profiles[profileId] = profile;
    }
  }
}
