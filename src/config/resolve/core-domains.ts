import { ResolveContext } from './context';
import { applyControllerConfig } from './controller';
import { isNotificationType, isOverlayNotificationPosition } from '../../types/notification';
import { asBoolean, asNumber, asString, isObject } from './shared';

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

    const logRotation = src.logging.rotation;
    if (typeof logRotation === 'number' && Number.isInteger(logRotation) && logRotation > 0) {
      resolved.logging.rotation = logRotation;
    } else if (src.logging.rotation !== undefined) {
      warn(
        'logging.rotation',
        src.logging.rotation,
        resolved.logging.rotation,
        'Expected a positive whole number of days.',
      );
    }

    if (isObject(src.logging.files)) {
      for (const key of ['app', 'launcher', 'mpv'] as const) {
        const enabled = asBoolean(src.logging.files[key]);
        if (enabled !== undefined) {
          resolved.logging.files[key] = enabled;
        } else if (src.logging.files[key] !== undefined) {
          warn(
            `logging.files.${key}`,
            src.logging.files[key],
            resolved.logging.files[key],
            'Expected boolean.',
          );
        }
      }
    } else if (src.logging.files !== undefined) {
      warn('logging.files', src.logging.files, resolved.logging.files, 'Expected object.');
    }
  }

  applyControllerConfig(context);

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

  if (isObject(src.updates)) {
    const enabled = asBoolean(src.updates.enabled);
    if (enabled !== undefined) {
      resolved.updates.enabled = enabled;
    } else if (src.updates.enabled !== undefined) {
      warn('updates.enabled', src.updates.enabled, resolved.updates.enabled, 'Expected boolean.');
    }

    const checkIntervalHours = asNumber(src.updates.checkIntervalHours);
    if (
      checkIntervalHours !== undefined &&
      Number.isFinite(checkIntervalHours) &&
      checkIntervalHours > 0
    ) {
      resolved.updates.checkIntervalHours = checkIntervalHours;
    } else if (src.updates.checkIntervalHours !== undefined) {
      warn(
        'updates.checkIntervalHours',
        src.updates.checkIntervalHours,
        resolved.updates.checkIntervalHours,
        'Expected positive number.',
      );
    }

    const notificationType = asString(src.updates.notificationType);
    if (isNotificationType(notificationType)) {
      resolved.updates.notificationType = notificationType;
    } else if (src.updates.notificationType !== undefined) {
      warn(
        'updates.notificationType',
        src.updates.notificationType,
        resolved.updates.notificationType,
        'Expected overlay, system, both, none, osd, or osd-system.',
      );
    }

    const channel = asString(src.updates.channel);
    if (channel === 'stable' || channel === 'prerelease') {
      resolved.updates.channel = channel;
    } else if (src.updates.channel !== undefined) {
      warn(
        'updates.channel',
        src.updates.channel,
        resolved.updates.channel,
        'Expected stable or prerelease.',
      );
    }
  } else if (src.updates !== undefined) {
    warn('updates', src.updates, resolved.updates, 'Expected object.');
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
      'openCharacterDictionaryManager',
      'openRuntimeOptions',
      'openJimaku',
      'toggleNotificationHistory',
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

  if (isObject(src.notifications)) {
    const overlayPosition = asString(src.notifications.overlayPosition);
    if (isOverlayNotificationPosition(overlayPosition)) {
      resolved.notifications.overlayPosition = overlayPosition;
    } else if (src.notifications.overlayPosition !== undefined) {
      warn(
        'notifications.overlayPosition',
        src.notifications.overlayPosition,
        resolved.notifications.overlayPosition,
        'Expected top-left, top, or top-right.',
      );
    }
  }
}
