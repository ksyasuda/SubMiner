import { ResolveContext } from './context';
import { asBoolean, asNumber, asString, isObject } from './shared';

export function applyCoreDomainConfig(context: ResolveContext): void {
  const { src, resolved, warn } = context;
  const controllerButtonBindings = [
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
  const controllerAxisBindings = [
    'leftStickX',
    'leftStickY',
    'rightStickX',
    'rightStickY',
  ] as const;

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
        const value = asString(src.controller.bindings[key]);
        if (
          value !== undefined &&
          controllerButtonBindings.includes(value as (typeof controllerButtonBindings)[number])
        ) {
          resolved.controller.bindings[key] =
            value as (typeof resolved.controller.bindings)[typeof key];
        } else if (src.controller.bindings[key] !== undefined) {
          warn(
            `controller.bindings.${key}`,
            src.controller.bindings[key],
            resolved.controller.bindings[key],
            `Expected one of: ${controllerButtonBindings.join(', ')}.`,
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
        const value = asString(src.controller.bindings[key]);
        if (
          value !== undefined &&
          controllerAxisBindings.includes(value as (typeof controllerAxisBindings)[number])
        ) {
          resolved.controller.bindings[key] =
            value as (typeof resolved.controller.bindings)[typeof key];
        } else if (src.controller.bindings[key] !== undefined) {
          warn(
            `controller.bindings.${key}`,
            src.controller.bindings[key],
            resolved.controller.bindings[key],
            `Expected one of: ${controllerAxisBindings.join(', ')}.`,
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
