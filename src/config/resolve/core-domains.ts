import type {
  ControllerAxisBinding,
  ControllerAxisBindingConfig,
  ControllerAxisDirection,
  ControllerButtonBinding,
  ControllerButtonIndicesConfig,
  ControllerDpadFallback,
  ControllerDiscreteBindingConfig,
  ResolvedControllerAxisBinding,
  ResolvedControllerDiscreteBinding,
} from '../../types';
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

const CONTROLLER_AXIS_BINDINGS = ['leftStickX', 'leftStickY', 'rightStickX', 'rightStickY'] as const;

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
    return typeof value.buttonIndex === 'number' && Number.isInteger(value.buttonIndex) && value.buttonIndex >= 0
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
  if (typeof value.axisIndex !== 'number' || !Number.isInteger(value.axisIndex) || value.axisIndex < 0) {
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

export function applyCoreDomainConfig(context: ResolveContext): void {
  const { src, resolved, warn } = context;

  if (isObject(src.texthooker)) {
    const launchAtStartup = asBoolean(src.texthooker.launchAtStartup);
    if (launchAtStartup !== undefined) {
      resolved.texthooker.launchAtStartup = launchAtStartup;
    } else if (src.texthooker.launchAtStartup !== undefined) {
      warn(
        'texthooker.launchAtStartup',
        src.texthooker.launchAtStartup,
        resolved.texthooker.launchAtStartup,
        'Expected boolean.',
      );
    }

    const openBrowser = asBoolean(src.texthooker.openBrowser);
    if (openBrowser !== undefined) {
      resolved.texthooker.openBrowser = openBrowser;
    } else if (src.texthooker.openBrowser !== undefined) {
      warn(
        'texthooker.openBrowser',
        src.texthooker.openBrowser,
        resolved.texthooker.openBrowser,
        'Expected boolean.',
      );
    }
  }

  if (isObject(src.websocket)) {
    const enabled = src.websocket.enabled;
    if (enabled === 'auto' || enabled === true || enabled === false) {
      resolved.websocket.enabled = enabled;
    } else if (enabled !== undefined) {
      warn(
        'websocket.enabled',
        enabled,
        resolved.websocket.enabled,
        "Expected true, false, or 'auto'.",
      );
    }

    const port = asNumber(src.websocket.port);
    if (port !== undefined && port > 0 && port <= 65535) {
      resolved.websocket.port = Math.floor(port);
    } else if (src.websocket.port !== undefined) {
      warn(
        'websocket.port',
        src.websocket.port,
        resolved.websocket.port,
        'Expected integer between 1 and 65535.',
      );
    }
  }

  if (isObject(src.annotationWebsocket)) {
    const enabled = asBoolean(src.annotationWebsocket.enabled);
    if (enabled !== undefined) {
      resolved.annotationWebsocket.enabled = enabled;
    } else if (src.annotationWebsocket.enabled !== undefined) {
      warn(
        'annotationWebsocket.enabled',
        src.annotationWebsocket.enabled,
        resolved.annotationWebsocket.enabled,
        'Expected boolean.',
      );
    }

    const port = asNumber(src.annotationWebsocket.port);
    if (port !== undefined && port > 0 && port <= 65535) {
      resolved.annotationWebsocket.port = Math.floor(port);
    } else if (src.annotationWebsocket.port !== undefined) {
      warn(
        'annotationWebsocket.port',
        src.annotationWebsocket.port,
        resolved.annotationWebsocket.port,
        'Expected integer between 1 and 65535.',
      );
    }
  }

  if (isObject(src.logging)) {
    const logLevel = asString(src.logging.level);
    if (
      logLevel === 'debug' ||
      logLevel === 'info' ||
      logLevel === 'warn' ||
      logLevel === 'error'
    ) {
      resolved.logging.level = logLevel;
    } else if (src.logging.level !== undefined) {
      warn(
        'logging.level',
        src.logging.level,
        resolved.logging.level,
        'Expected debug, info, warn, or error.',
      );
    }
  }

  if (isObject(src.controller)) {
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

    if (isObject(src.controller.buttonIndices)) {
      const buttonIndexKeys = [
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

      for (const key of buttonIndexKeys) {
        const value = asNumber(src.controller.buttonIndices[key]);
        if (value !== undefined && value >= 0 && Number.isInteger(value)) {
          resolved.controller.buttonIndices[key] = value;
        } else if (src.controller.buttonIndices[key] !== undefined) {
          warn(
            `controller.buttonIndices.${key}`,
            src.controller.buttonIndices[key],
            resolved.controller.buttonIndices[key],
            'Expected non-negative integer.',
          );
        }
      }
    }

    if (isObject(src.controller.bindings)) {
      const buttonBindingKeys = [
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

      for (const key of buttonBindingKeys) {
        const bindingValue = src.controller.bindings[key];
        const legacyValue = asString(bindingValue);
        if (
          legacyValue !== undefined &&
          CONTROLLER_BUTTON_BINDINGS.includes(legacyValue as (typeof CONTROLLER_BUTTON_BINDINGS)[number])
        ) {
          resolved.controller.bindings[key] = resolveLegacyDiscreteBinding(
            legacyValue as ControllerButtonBinding,
            resolved.controller.buttonIndices,
          );
          continue;
        }
        const parsedObject = parseDiscreteBindingObject(bindingValue);
        if (parsedObject) {
          resolved.controller.bindings[key] = parsedObject;
        } else if (bindingValue !== undefined) {
          warn(
            `controller.bindings.${key}`,
            bindingValue,
            resolved.controller.bindings[key],
            "Expected legacy controller button name or binding object with kind 'none', 'button', or 'axis'.",
          );
        }
      }

      const axisBindingKeys = [
        'leftStickHorizontal',
        'leftStickVertical',
        'rightStickHorizontal',
        'rightStickVertical',
      ] as const;

      for (const key of axisBindingKeys) {
        const bindingValue = src.controller.bindings[key];
        const legacyValue = asString(bindingValue);
        if (
          legacyValue !== undefined &&
          CONTROLLER_AXIS_BINDINGS.includes(legacyValue as (typeof CONTROLLER_AXIS_BINDINGS)[number])
        ) {
          resolved.controller.bindings[key] = resolveLegacyAxisBinding(
            legacyValue as ControllerAxisBinding,
            key,
          );
          continue;
        }
        if (legacyValue === 'none') {
          resolved.controller.bindings[key] = { kind: 'none' };
          continue;
        }
        const parsedObject = parseAxisBindingObject(bindingValue, key);
        if (parsedObject) {
          resolved.controller.bindings[key] = parsedObject;
        } else if (bindingValue !== undefined) {
          warn(
            `controller.bindings.${key}`,
            bindingValue,
            resolved.controller.bindings[key],
            "Expected legacy controller axis name ('none' allowed) or binding object with kind 'axis'.",
          );
        }
      }
    }
  }

  if (Array.isArray(src.keybindings)) {
    resolved.keybindings = src.keybindings.filter(
      (entry): entry is { key: string; command: (string | number)[] | null } => {
        if (!isObject(entry)) return false;
        if (typeof entry.key !== 'string') return false;
        if (entry.command === null) return true;
        return Array.isArray(entry.command);
      },
    );
  }

  if (isObject(src.startupWarmups)) {
    const startupWarmupBooleanKeys = [
      'lowPowerMode',
      'mecab',
      'yomitanExtension',
      'subtitleDictionaries',
      'jellyfinRemoteSession',
    ] as const;

    for (const key of startupWarmupBooleanKeys) {
      const value = asBoolean(src.startupWarmups[key]);
      if (value !== undefined) {
        resolved.startupWarmups[key] = value as (typeof resolved.startupWarmups)[typeof key];
      } else if (src.startupWarmups[key] !== undefined) {
        warn(
          `startupWarmups.${key}`,
          src.startupWarmups[key],
          resolved.startupWarmups[key],
          'Expected boolean.',
        );
      }
    }
  }

  if (isObject(src.shortcuts)) {
    const shortcutKeys = [
      'toggleVisibleOverlayGlobal',
      'copySubtitle',
      'copySubtitleMultiple',
      'updateLastCardFromClipboard',
      'triggerFieldGrouping',
      'triggerSubsync',
      'mineSentence',
      'mineSentenceMultiple',
      'toggleSecondarySub',
      'markAudioCard',
      'openRuntimeOptions',
      'openJimaku',
    ] as const;

    for (const key of shortcutKeys) {
      const value = src.shortcuts[key];
      if (typeof value === 'string' || value === null) {
        resolved.shortcuts[key] = value as (typeof resolved.shortcuts)[typeof key];
      } else if (value !== undefined) {
        warn(`shortcuts.${key}`, value, resolved.shortcuts[key], 'Expected string or null.');
      }
    }

    const timeout = asNumber(src.shortcuts.multiCopyTimeoutMs);
    if (timeout !== undefined && timeout > 0) {
      resolved.shortcuts.multiCopyTimeoutMs = Math.floor(timeout);
    } else if (src.shortcuts.multiCopyTimeoutMs !== undefined) {
      warn(
        'shortcuts.multiCopyTimeoutMs',
        src.shortcuts.multiCopyTimeoutMs,
        resolved.shortcuts.multiCopyTimeoutMs,
        'Expected positive number.',
      );
    }
  }

  if (isObject(src.secondarySub)) {
    if (Array.isArray(src.secondarySub.secondarySubLanguages)) {
      resolved.secondarySub.secondarySubLanguages = src.secondarySub.secondarySubLanguages.filter(
        (item): item is string => typeof item === 'string',
      );
    }
    const autoLoad = asBoolean(src.secondarySub.autoLoadSecondarySub);
    if (autoLoad !== undefined) {
      resolved.secondarySub.autoLoadSecondarySub = autoLoad;
    }
    const defaultMode = src.secondarySub.defaultMode;
    if (defaultMode === 'hidden' || defaultMode === 'visible' || defaultMode === 'hover') {
      resolved.secondarySub.defaultMode = defaultMode;
    } else if (defaultMode !== undefined) {
      warn(
        'secondarySub.defaultMode',
        defaultMode,
        resolved.secondarySub.defaultMode,
        'Expected hidden, visible, or hover.',
      );
    }
  }

  if (isObject(src.youtube)) {
    if (Array.isArray(src.youtube.primarySubLanguages)) {
      resolved.youtube.primarySubLanguages = src.youtube.primarySubLanguages.filter(
        (item): item is string => typeof item === 'string',
      );
    } else if (src.youtube.primarySubLanguages !== undefined) {
      warn(
        'youtube.primarySubLanguages',
        src.youtube.primarySubLanguages,
        resolved.youtube.primarySubLanguages,
        'Expected string array.',
      );
    }
  }

  if (isObject(src.subsync)) {
    const mode = src.subsync.defaultMode;
    if (mode === 'auto' || mode === 'manual') {
      resolved.subsync.defaultMode = mode;
    } else if (mode !== undefined) {
      warn('subsync.defaultMode', mode, resolved.subsync.defaultMode, 'Expected auto or manual.');
    }

    const alass = asString(src.subsync.alass_path);
    if (alass !== undefined) resolved.subsync.alass_path = alass;
    const ffsubsync = asString(src.subsync.ffsubsync_path);
    if (ffsubsync !== undefined) resolved.subsync.ffsubsync_path = ffsubsync;
    const ffmpeg = asString(src.subsync.ffmpeg_path);
    if (ffmpeg !== undefined) resolved.subsync.ffmpeg_path = ffmpeg;
    const replace = asBoolean(src.subsync.replace);
    if (replace !== undefined) {
      resolved.subsync.replace = replace;
    } else if (src.subsync.replace !== undefined) {
      warn('subsync.replace', src.subsync.replace, resolved.subsync.replace, 'Expected boolean.');
    }
  }

  if (isObject(src.subtitlePosition)) {
    const y = asNumber(src.subtitlePosition.yPercent);
    if (y !== undefined) {
      resolved.subtitlePosition.yPercent = y;
    }
  }
}
