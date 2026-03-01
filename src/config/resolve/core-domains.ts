import { ResolveContext } from './context';
import { asBoolean, asNumber, asString, isObject } from './shared';

export function applyCoreDomainConfig(context: ResolveContext): void {
  const { src, resolved, warn } = context;

  if (isObject(src.texthooker)) {
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
  }

  if (isObject(src.subtitlePosition)) {
    const y = asNumber(src.subtitlePosition.yPercent);
    if (y !== undefined) {
      resolved.subtitlePosition.yPercent = y;
    }
  }
}
