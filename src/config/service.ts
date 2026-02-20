import * as fs from 'fs';
import * as path from 'path';
import { parse as parseJsonc, type ParseError } from 'jsonc-parser';
import { Config, ConfigValidationWarning, RawConfig, ResolvedConfig } from '../types';
import { DEFAULT_CONFIG, deepCloneConfig, deepMergeRawConfig } from './definitions';

interface LoadResult {
  config: RawConfig;
  path: string;
}

export type ReloadConfigStrictResult =
  | {
      ok: true;
      config: ResolvedConfig;
      warnings: ConfigValidationWarning[];
      path: string;
    }
  | {
      ok: false;
      error: string;
      path: string;
    };

function isObject(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function asNumber(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) ? value : undefined;
}

function asString(value: unknown): string | undefined {
  return typeof value === 'string' ? value : undefined;
}

function asBoolean(value: unknown): boolean | undefined {
  return typeof value === 'boolean' ? value : undefined;
}

const hexColorPattern = /^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/;

function asColor(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined;
  const text = value.trim();
  return hexColorPattern.test(text) ? text : undefined;
}

function asFrequencyBandedColors(
  value: unknown,
): [string, string, string, string, string] | undefined {
  if (!Array.isArray(value) || value.length !== 5) {
    return undefined;
  }

  const colors = value.map((item) => asColor(item));
  if (colors.some((color) => color === undefined)) {
    return undefined;
  }

  return colors as [string, string, string, string, string];
}

export class ConfigService {
  private readonly configDir: string;
  private readonly configFileJsonc: string;
  private readonly configFileJson: string;
  private rawConfig: RawConfig = {};
  private resolvedConfig: ResolvedConfig = deepCloneConfig(DEFAULT_CONFIG);
  private warnings: ConfigValidationWarning[] = [];
  private configPathInUse: string;

  constructor(configDir: string) {
    this.configDir = configDir;
    this.configFileJsonc = path.join(configDir, 'config.jsonc');
    this.configFileJson = path.join(configDir, 'config.json');
    this.configPathInUse = this.configFileJsonc;
    this.reloadConfig();
  }

  getConfigPath(): string {
    return this.configPathInUse;
  }

  getConfig(): ResolvedConfig {
    return deepCloneConfig(this.resolvedConfig);
  }

  getRawConfig(): RawConfig {
    return JSON.parse(JSON.stringify(this.rawConfig)) as RawConfig;
  }

  getWarnings(): ConfigValidationWarning[] {
    return [...this.warnings];
  }

  reloadConfig(): ResolvedConfig {
    const { config, path: configPath } = this.loadRawConfig();
    this.rawConfig = config;
    this.configPathInUse = configPath;
    const { resolved, warnings } = this.resolveConfig(config);
    this.resolvedConfig = resolved;
    this.warnings = warnings;
    return this.getConfig();
  }

  reloadConfigStrict(): ReloadConfigStrictResult {
    const loadResult = this.loadRawConfigStrict();
    if (!loadResult.ok) {
      return loadResult;
    }

    const { config, path: configPath } = loadResult;
    this.rawConfig = config;
    this.configPathInUse = configPath;
    const { resolved, warnings } = this.resolveConfig(config);
    this.resolvedConfig = resolved;
    this.warnings = warnings;
    return {
      ok: true,
      config: this.getConfig(),
      warnings: [...warnings],
      path: configPath,
    };
  }

  saveRawConfig(config: RawConfig): void {
    if (!fs.existsSync(this.configDir)) {
      fs.mkdirSync(this.configDir, { recursive: true });
    }
    const targetPath = this.configPathInUse.endsWith('.json')
      ? this.configPathInUse
      : this.configFileJsonc;
    fs.writeFileSync(targetPath, JSON.stringify(config, null, 2));
    this.rawConfig = config;
    this.configPathInUse = targetPath;
    const { resolved, warnings } = this.resolveConfig(config);
    this.resolvedConfig = resolved;
    this.warnings = warnings;
  }

  patchRawConfig(patch: RawConfig): void {
    const merged = deepMergeRawConfig(this.getRawConfig(), patch);
    this.saveRawConfig(merged);
  }

  private loadRawConfig(): LoadResult {
    const strictResult = this.loadRawConfigStrict();
    if (strictResult.ok) {
      return strictResult;
    }
    return { config: {}, path: strictResult.path };
  }

  private loadRawConfigStrict():
    | (LoadResult & { ok: true })
    | {
        ok: false;
        error: string;
        path: string;
      } {
    const configPath = fs.existsSync(this.configFileJsonc)
      ? this.configFileJsonc
      : fs.existsSync(this.configFileJson)
        ? this.configFileJson
        : this.configFileJsonc;

    if (!fs.existsSync(configPath)) {
      return { ok: true, config: {}, path: configPath };
    }

    try {
      const data = fs.readFileSync(configPath, 'utf-8');
      const parsed = configPath.endsWith('.jsonc')
        ? (() => {
            const errors: ParseError[] = [];
            const result = parseJsonc(data, errors, {
              allowTrailingComma: true,
              disallowComments: false,
            });
            if (errors.length > 0) {
              throw new Error(`Invalid JSONC (${errors[0]?.error ?? 'unknown'})`);
            }
            return result;
          })()
        : JSON.parse(data);
      return {
        ok: true,
        config: isObject(parsed) ? (parsed as Config) : {},
        path: configPath,
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown parse error';
      return { ok: false, error: message, path: configPath };
    }
  }

  private resolveConfig(raw: RawConfig): {
    resolved: ResolvedConfig;
    warnings: ConfigValidationWarning[];
  } {
    const warnings: ConfigValidationWarning[] = [];
    const resolved = deepCloneConfig(DEFAULT_CONFIG);

    const warn = (path: string, value: unknown, fallback: unknown, message: string): void => {
      warnings.push({
        path,
        value,
        fallback,
        message,
      });
    };

    const src = isObject(raw) ? raw : {};
    const knownTopLevelKeys = new Set(Object.keys(resolved));
    for (const key of Object.keys(src)) {
      if (!knownTopLevelKeys.has(key)) {
        warn(key, src[key], undefined, 'Unknown top-level config key; ignored.');
      }
    }

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

    if (isObject(src.shortcuts)) {
      const shortcutKeys = [
        'toggleVisibleOverlayGlobal',
        'toggleInvisibleOverlayGlobal',
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

    if (isObject(src.invisibleOverlay)) {
      const startupVisibility = src.invisibleOverlay.startupVisibility;
      if (
        startupVisibility === 'platform-default' ||
        startupVisibility === 'visible' ||
        startupVisibility === 'hidden'
      ) {
        resolved.invisibleOverlay.startupVisibility = startupVisibility;
      } else if (startupVisibility !== undefined) {
        warn(
          'invisibleOverlay.startupVisibility',
          startupVisibility,
          resolved.invisibleOverlay.startupVisibility,
          'Expected platform-default, visible, or hidden.',
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

    if (isObject(src.jimaku)) {
      const apiKey = asString(src.jimaku.apiKey);
      if (apiKey !== undefined) resolved.jimaku.apiKey = apiKey;
      const apiKeyCommand = asString(src.jimaku.apiKeyCommand);
      if (apiKeyCommand !== undefined) resolved.jimaku.apiKeyCommand = apiKeyCommand;
      const apiBaseUrl = asString(src.jimaku.apiBaseUrl);
      if (apiBaseUrl !== undefined) resolved.jimaku.apiBaseUrl = apiBaseUrl;

      const lang = src.jimaku.languagePreference;
      if (lang === 'ja' || lang === 'en' || lang === 'none') {
        resolved.jimaku.languagePreference = lang;
      } else if (lang !== undefined) {
        warn(
          'jimaku.languagePreference',
          lang,
          resolved.jimaku.languagePreference,
          'Expected ja, en, or none.',
        );
      }

      const maxEntryResults = asNumber(src.jimaku.maxEntryResults);
      if (maxEntryResults !== undefined && maxEntryResults > 0) {
        resolved.jimaku.maxEntryResults = Math.floor(maxEntryResults);
      } else if (src.jimaku.maxEntryResults !== undefined) {
        warn(
          'jimaku.maxEntryResults',
          src.jimaku.maxEntryResults,
          resolved.jimaku.maxEntryResults,
          'Expected positive number.',
        );
      }
    }

    if (isObject(src.youtubeSubgen)) {
      const mode = src.youtubeSubgen.mode;
      if (mode === 'automatic' || mode === 'preprocess' || mode === 'off') {
        resolved.youtubeSubgen.mode = mode;
      } else if (mode !== undefined) {
        warn(
          'youtubeSubgen.mode',
          mode,
          resolved.youtubeSubgen.mode,
          'Expected automatic, preprocess, or off.',
        );
      }

      const whisperBin = asString(src.youtubeSubgen.whisperBin);
      if (whisperBin !== undefined) {
        resolved.youtubeSubgen.whisperBin = whisperBin;
      } else if (src.youtubeSubgen.whisperBin !== undefined) {
        warn(
          'youtubeSubgen.whisperBin',
          src.youtubeSubgen.whisperBin,
          resolved.youtubeSubgen.whisperBin,
          'Expected string.',
        );
      }

      const whisperModel = asString(src.youtubeSubgen.whisperModel);
      if (whisperModel !== undefined) {
        resolved.youtubeSubgen.whisperModel = whisperModel;
      } else if (src.youtubeSubgen.whisperModel !== undefined) {
        warn(
          'youtubeSubgen.whisperModel',
          src.youtubeSubgen.whisperModel,
          resolved.youtubeSubgen.whisperModel,
          'Expected string.',
        );
      }

      if (Array.isArray(src.youtubeSubgen.primarySubLanguages)) {
        resolved.youtubeSubgen.primarySubLanguages = src.youtubeSubgen.primarySubLanguages.filter(
          (item): item is string => typeof item === 'string',
        );
      } else if (src.youtubeSubgen.primarySubLanguages !== undefined) {
        warn(
          'youtubeSubgen.primarySubLanguages',
          src.youtubeSubgen.primarySubLanguages,
          resolved.youtubeSubgen.primarySubLanguages,
          'Expected string array.',
        );
      }
    }

    if (isObject(src.anilist)) {
      const enabled = asBoolean(src.anilist.enabled);
      if (enabled !== undefined) {
        resolved.anilist.enabled = enabled;
      } else if (src.anilist.enabled !== undefined) {
        warn('anilist.enabled', src.anilist.enabled, resolved.anilist.enabled, 'Expected boolean.');
      }

      const accessToken = asString(src.anilist.accessToken);
      if (accessToken !== undefined) {
        resolved.anilist.accessToken = accessToken;
      } else if (src.anilist.accessToken !== undefined) {
        warn(
          'anilist.accessToken',
          src.anilist.accessToken,
          resolved.anilist.accessToken,
          'Expected string.',
        );
      }
    }

    if (isObject(src.jellyfin)) {
      const enabled = asBoolean(src.jellyfin.enabled);
      if (enabled !== undefined) {
        resolved.jellyfin.enabled = enabled;
      } else if (src.jellyfin.enabled !== undefined) {
        warn(
          'jellyfin.enabled',
          src.jellyfin.enabled,
          resolved.jellyfin.enabled,
          'Expected boolean.',
        );
      }

      const stringKeys = [
        'serverUrl',
        'username',
        'accessToken',
        'userId',
        'deviceId',
        'clientName',
        'clientVersion',
        'defaultLibraryId',
        'iconCacheDir',
        'transcodeVideoCodec',
      ] as const;
      for (const key of stringKeys) {
        const value = asString(src.jellyfin[key]);
        if (value !== undefined) {
          resolved.jellyfin[key] = value as (typeof resolved.jellyfin)[typeof key];
        } else if (src.jellyfin[key] !== undefined) {
          warn(`jellyfin.${key}`, src.jellyfin[key], resolved.jellyfin[key], 'Expected string.');
        }
      }

      const booleanKeys = [
        'remoteControlEnabled',
        'remoteControlAutoConnect',
        'autoAnnounce',
        'directPlayPreferred',
        'pullPictures',
      ] as const;
      for (const key of booleanKeys) {
        const value = asBoolean(src.jellyfin[key]);
        if (value !== undefined) {
          resolved.jellyfin[key] = value as (typeof resolved.jellyfin)[typeof key];
        } else if (src.jellyfin[key] !== undefined) {
          warn(`jellyfin.${key}`, src.jellyfin[key], resolved.jellyfin[key], 'Expected boolean.');
        }
      }

      if (Array.isArray(src.jellyfin.directPlayContainers)) {
        resolved.jellyfin.directPlayContainers = src.jellyfin.directPlayContainers
          .filter((item): item is string => typeof item === 'string')
          .map((item) => item.trim().toLowerCase())
          .filter((item) => item.length > 0);
      } else if (src.jellyfin.directPlayContainers !== undefined) {
        warn(
          'jellyfin.directPlayContainers',
          src.jellyfin.directPlayContainers,
          resolved.jellyfin.directPlayContainers,
          'Expected string array.',
        );
      }
    }

    if (asBoolean(src.auto_start_overlay) !== undefined) {
      resolved.auto_start_overlay = src.auto_start_overlay as boolean;
    }

    if (asBoolean(src.bind_visible_overlay_to_mpv_sub_visibility) !== undefined) {
      resolved.bind_visible_overlay_to_mpv_sub_visibility =
        src.bind_visible_overlay_to_mpv_sub_visibility as boolean;
    } else if (src.bind_visible_overlay_to_mpv_sub_visibility !== undefined) {
      warn(
        'bind_visible_overlay_to_mpv_sub_visibility',
        src.bind_visible_overlay_to_mpv_sub_visibility,
        resolved.bind_visible_overlay_to_mpv_sub_visibility,
        'Expected boolean.',
      );
    }

    if (isObject(src.immersionTracking)) {
      const enabled = asBoolean(src.immersionTracking.enabled);
      if (enabled !== undefined) {
        resolved.immersionTracking.enabled = enabled;
      } else if (src.immersionTracking.enabled !== undefined) {
        warn(
          'immersionTracking.enabled',
          src.immersionTracking.enabled,
          resolved.immersionTracking.enabled,
          'Expected boolean.',
        );
      }

      const dbPath = asString(src.immersionTracking.dbPath);
      if (dbPath !== undefined) {
        resolved.immersionTracking.dbPath = dbPath;
      } else if (src.immersionTracking.dbPath !== undefined) {
        warn(
          'immersionTracking.dbPath',
          src.immersionTracking.dbPath,
          resolved.immersionTracking.dbPath,
          'Expected string.',
        );
      }

      const batchSize = asNumber(src.immersionTracking.batchSize);
      if (batchSize !== undefined && batchSize >= 1 && batchSize <= 10_000) {
        resolved.immersionTracking.batchSize = Math.floor(batchSize);
      } else if (src.immersionTracking.batchSize !== undefined) {
        warn(
          'immersionTracking.batchSize',
          src.immersionTracking.batchSize,
          resolved.immersionTracking.batchSize,
          'Expected integer between 1 and 10000.',
        );
      }

      const flushIntervalMs = asNumber(src.immersionTracking.flushIntervalMs);
      if (flushIntervalMs !== undefined && flushIntervalMs >= 50 && flushIntervalMs <= 60_000) {
        resolved.immersionTracking.flushIntervalMs = Math.floor(flushIntervalMs);
      } else if (src.immersionTracking.flushIntervalMs !== undefined) {
        warn(
          'immersionTracking.flushIntervalMs',
          src.immersionTracking.flushIntervalMs,
          resolved.immersionTracking.flushIntervalMs,
          'Expected integer between 50 and 60000.',
        );
      }

      const queueCap = asNumber(src.immersionTracking.queueCap);
      if (queueCap !== undefined && queueCap >= 100 && queueCap <= 100_000) {
        resolved.immersionTracking.queueCap = Math.floor(queueCap);
      } else if (src.immersionTracking.queueCap !== undefined) {
        warn(
          'immersionTracking.queueCap',
          src.immersionTracking.queueCap,
          resolved.immersionTracking.queueCap,
          'Expected integer between 100 and 100000.',
        );
      }

      const payloadCapBytes = asNumber(src.immersionTracking.payloadCapBytes);
      if (payloadCapBytes !== undefined && payloadCapBytes >= 64 && payloadCapBytes <= 8192) {
        resolved.immersionTracking.payloadCapBytes = Math.floor(payloadCapBytes);
      } else if (src.immersionTracking.payloadCapBytes !== undefined) {
        warn(
          'immersionTracking.payloadCapBytes',
          src.immersionTracking.payloadCapBytes,
          resolved.immersionTracking.payloadCapBytes,
          'Expected integer between 64 and 8192.',
        );
      }

      const maintenanceIntervalMs = asNumber(src.immersionTracking.maintenanceIntervalMs);
      if (
        maintenanceIntervalMs !== undefined &&
        maintenanceIntervalMs >= 60_000 &&
        maintenanceIntervalMs <= 7 * 24 * 60 * 60 * 1000
      ) {
        resolved.immersionTracking.maintenanceIntervalMs = Math.floor(maintenanceIntervalMs);
      } else if (src.immersionTracking.maintenanceIntervalMs !== undefined) {
        warn(
          'immersionTracking.maintenanceIntervalMs',
          src.immersionTracking.maintenanceIntervalMs,
          resolved.immersionTracking.maintenanceIntervalMs,
          'Expected integer between 60000 and 604800000.',
        );
      }

      if (isObject(src.immersionTracking.retention)) {
        const eventsDays = asNumber(src.immersionTracking.retention.eventsDays);
        if (eventsDays !== undefined && eventsDays >= 1 && eventsDays <= 3650) {
          resolved.immersionTracking.retention.eventsDays = Math.floor(eventsDays);
        } else if (src.immersionTracking.retention.eventsDays !== undefined) {
          warn(
            'immersionTracking.retention.eventsDays',
            src.immersionTracking.retention.eventsDays,
            resolved.immersionTracking.retention.eventsDays,
            'Expected integer between 1 and 3650.',
          );
        }

        const telemetryDays = asNumber(src.immersionTracking.retention.telemetryDays);
        if (telemetryDays !== undefined && telemetryDays >= 1 && telemetryDays <= 3650) {
          resolved.immersionTracking.retention.telemetryDays = Math.floor(telemetryDays);
        } else if (src.immersionTracking.retention.telemetryDays !== undefined) {
          warn(
            'immersionTracking.retention.telemetryDays',
            src.immersionTracking.retention.telemetryDays,
            resolved.immersionTracking.retention.telemetryDays,
            'Expected integer between 1 and 3650.',
          );
        }

        const dailyRollupsDays = asNumber(src.immersionTracking.retention.dailyRollupsDays);
        if (dailyRollupsDays !== undefined && dailyRollupsDays >= 1 && dailyRollupsDays <= 36500) {
          resolved.immersionTracking.retention.dailyRollupsDays = Math.floor(dailyRollupsDays);
        } else if (src.immersionTracking.retention.dailyRollupsDays !== undefined) {
          warn(
            'immersionTracking.retention.dailyRollupsDays',
            src.immersionTracking.retention.dailyRollupsDays,
            resolved.immersionTracking.retention.dailyRollupsDays,
            'Expected integer between 1 and 36500.',
          );
        }

        const monthlyRollupsDays = asNumber(src.immersionTracking.retention.monthlyRollupsDays);
        if (
          monthlyRollupsDays !== undefined &&
          monthlyRollupsDays >= 1 &&
          monthlyRollupsDays <= 36500
        ) {
          resolved.immersionTracking.retention.monthlyRollupsDays = Math.floor(monthlyRollupsDays);
        } else if (src.immersionTracking.retention.monthlyRollupsDays !== undefined) {
          warn(
            'immersionTracking.retention.monthlyRollupsDays',
            src.immersionTracking.retention.monthlyRollupsDays,
            resolved.immersionTracking.retention.monthlyRollupsDays,
            'Expected integer between 1 and 36500.',
          );
        }

        const vacuumIntervalDays = asNumber(src.immersionTracking.retention.vacuumIntervalDays);
        if (
          vacuumIntervalDays !== undefined &&
          vacuumIntervalDays >= 1 &&
          vacuumIntervalDays <= 3650
        ) {
          resolved.immersionTracking.retention.vacuumIntervalDays = Math.floor(vacuumIntervalDays);
        } else if (src.immersionTracking.retention.vacuumIntervalDays !== undefined) {
          warn(
            'immersionTracking.retention.vacuumIntervalDays',
            src.immersionTracking.retention.vacuumIntervalDays,
            resolved.immersionTracking.retention.vacuumIntervalDays,
            'Expected integer between 1 and 3650.',
          );
        }
      } else if (src.immersionTracking.retention !== undefined) {
        warn(
          'immersionTracking.retention',
          src.immersionTracking.retention,
          resolved.immersionTracking.retention,
          'Expected object.',
        );
      }
    }

    if (isObject(src.subtitleStyle)) {
      const fallbackSubtitleStyleEnableJlpt = resolved.subtitleStyle.enableJlpt;
      const fallbackSubtitleStylePreserveLineBreaks = resolved.subtitleStyle.preserveLineBreaks;
      resolved.subtitleStyle = {
        ...resolved.subtitleStyle,
        ...(src.subtitleStyle as ResolvedConfig['subtitleStyle']),
        secondary: {
          ...resolved.subtitleStyle.secondary,
          ...(isObject(src.subtitleStyle.secondary)
            ? (src.subtitleStyle.secondary as ResolvedConfig['subtitleStyle']['secondary'])
            : {}),
        },
      };

      const enableJlpt = asBoolean((src.subtitleStyle as { enableJlpt?: unknown }).enableJlpt);
      if (enableJlpt !== undefined) {
        resolved.subtitleStyle.enableJlpt = enableJlpt;
      } else if ((src.subtitleStyle as { enableJlpt?: unknown }).enableJlpt !== undefined) {
        resolved.subtitleStyle.enableJlpt = fallbackSubtitleStyleEnableJlpt;
        warn(
          'subtitleStyle.enableJlpt',
          (src.subtitleStyle as { enableJlpt?: unknown }).enableJlpt,
          resolved.subtitleStyle.enableJlpt,
          'Expected boolean.',
        );
      }

      const preserveLineBreaks = asBoolean(
        (src.subtitleStyle as { preserveLineBreaks?: unknown }).preserveLineBreaks,
      );
      if (preserveLineBreaks !== undefined) {
        resolved.subtitleStyle.preserveLineBreaks = preserveLineBreaks;
      } else if (
        (src.subtitleStyle as { preserveLineBreaks?: unknown }).preserveLineBreaks !== undefined
      ) {
        resolved.subtitleStyle.preserveLineBreaks = fallbackSubtitleStylePreserveLineBreaks;
        warn(
          'subtitleStyle.preserveLineBreaks',
          (src.subtitleStyle as { preserveLineBreaks?: unknown }).preserveLineBreaks,
          resolved.subtitleStyle.preserveLineBreaks,
          'Expected boolean.',
        );
      }

      const frequencyDictionary = isObject(
        (src.subtitleStyle as { frequencyDictionary?: unknown }).frequencyDictionary,
      )
        ? ((src.subtitleStyle as { frequencyDictionary?: unknown }).frequencyDictionary as Record<
            string,
            unknown
          >)
        : {};
      const frequencyEnabled = asBoolean((frequencyDictionary as { enabled?: unknown }).enabled);
      if (frequencyEnabled !== undefined) {
        resolved.subtitleStyle.frequencyDictionary.enabled = frequencyEnabled;
      } else if ((frequencyDictionary as { enabled?: unknown }).enabled !== undefined) {
        warn(
          'subtitleStyle.frequencyDictionary.enabled',
          (frequencyDictionary as { enabled?: unknown }).enabled,
          resolved.subtitleStyle.frequencyDictionary.enabled,
          'Expected boolean.',
        );
      }

      const sourcePath = asString((frequencyDictionary as { sourcePath?: unknown }).sourcePath);
      if (sourcePath !== undefined) {
        resolved.subtitleStyle.frequencyDictionary.sourcePath = sourcePath;
      } else if ((frequencyDictionary as { sourcePath?: unknown }).sourcePath !== undefined) {
        warn(
          'subtitleStyle.frequencyDictionary.sourcePath',
          (frequencyDictionary as { sourcePath?: unknown }).sourcePath,
          resolved.subtitleStyle.frequencyDictionary.sourcePath,
          'Expected string.',
        );
      }

      const topX = asNumber((frequencyDictionary as { topX?: unknown }).topX);
      if (topX !== undefined && Number.isInteger(topX) && topX > 0) {
        resolved.subtitleStyle.frequencyDictionary.topX = Math.floor(topX);
      } else if ((frequencyDictionary as { topX?: unknown }).topX !== undefined) {
        warn(
          'subtitleStyle.frequencyDictionary.topX',
          (frequencyDictionary as { topX?: unknown }).topX,
          resolved.subtitleStyle.frequencyDictionary.topX,
          'Expected a positive integer.',
        );
      }

      const frequencyMode = frequencyDictionary.mode;
      if (frequencyMode === 'single' || frequencyMode === 'banded') {
        resolved.subtitleStyle.frequencyDictionary.mode = frequencyMode;
      } else if (frequencyMode !== undefined) {
        warn(
          'subtitleStyle.frequencyDictionary.mode',
          frequencyDictionary.mode,
          resolved.subtitleStyle.frequencyDictionary.mode,
          "Expected 'single' or 'banded'.",
        );
      }

      const singleColor = asColor((frequencyDictionary as { singleColor?: unknown }).singleColor);
      if (singleColor !== undefined) {
        resolved.subtitleStyle.frequencyDictionary.singleColor = singleColor;
      } else if ((frequencyDictionary as { singleColor?: unknown }).singleColor !== undefined) {
        warn(
          'subtitleStyle.frequencyDictionary.singleColor',
          (frequencyDictionary as { singleColor?: unknown }).singleColor,
          resolved.subtitleStyle.frequencyDictionary.singleColor,
          'Expected hex color.',
        );
      }

      const bandedColors = asFrequencyBandedColors(
        (frequencyDictionary as { bandedColors?: unknown }).bandedColors,
      );
      if (bandedColors !== undefined) {
        resolved.subtitleStyle.frequencyDictionary.bandedColors = bandedColors;
      } else if ((frequencyDictionary as { bandedColors?: unknown }).bandedColors !== undefined) {
        warn(
          'subtitleStyle.frequencyDictionary.bandedColors',
          (frequencyDictionary as { bandedColors?: unknown }).bandedColors,
          resolved.subtitleStyle.frequencyDictionary.bandedColors,
          'Expected an array of five hex colors.',
        );
      }
    }

    if (isObject(src.ankiConnect)) {
      const ac = src.ankiConnect;
      const behavior = isObject(ac.behavior) ? (ac.behavior as Record<string, unknown>) : {};
      const fields = isObject(ac.fields) ? (ac.fields as Record<string, unknown>) : {};
      const media = isObject(ac.media) ? (ac.media as Record<string, unknown>) : {};
      const metadata = isObject(ac.metadata) ? (ac.metadata as Record<string, unknown>) : {};
      const aiSource = isObject(ac.ai) ? ac.ai : isObject(ac.openRouter) ? ac.openRouter : {};
      const legacyKeys = new Set([
        'audioField',
        'imageField',
        'sentenceField',
        'miscInfoField',
        'miscInfoPattern',
        'generateAudio',
        'generateImage',
        'imageType',
        'imageFormat',
        'imageQuality',
        'imageMaxWidth',
        'imageMaxHeight',
        'animatedFps',
        'animatedMaxWidth',
        'animatedMaxHeight',
        'animatedCrf',
        'audioPadding',
        'fallbackDuration',
        'maxMediaDuration',
        'overwriteAudio',
        'overwriteImage',
        'mediaInsertMode',
        'highlightWord',
        'notificationType',
        'autoUpdateNewCards',
      ]);

      if (ac.openRouter !== undefined) {
        warn(
          'ankiConnect.openRouter',
          ac.openRouter,
          resolved.ankiConnect.ai,
          'Deprecated key; use ankiConnect.ai instead.',
        );
      }

      const { nPlusOne: _nPlusOneConfigFromAnkiConnect, ...ankiConnectWithoutNPlusOne } =
        ac as Record<string, unknown>;
      const ankiConnectWithoutLegacy = Object.fromEntries(
        Object.entries(ankiConnectWithoutNPlusOne).filter(([key]) => !legacyKeys.has(key)),
      );

      resolved.ankiConnect = {
        ...resolved.ankiConnect,
        ...(isObject(ankiConnectWithoutLegacy)
          ? (ankiConnectWithoutLegacy as Partial<ResolvedConfig['ankiConnect']>)
          : {}),
        fields: {
          ...resolved.ankiConnect.fields,
          ...(isObject(ac.fields) ? (ac.fields as ResolvedConfig['ankiConnect']['fields']) : {}),
        },
        ai: {
          ...resolved.ankiConnect.ai,
          ...(aiSource as ResolvedConfig['ankiConnect']['ai']),
        },
        media: {
          ...resolved.ankiConnect.media,
          ...(isObject(ac.media) ? (ac.media as ResolvedConfig['ankiConnect']['media']) : {}),
        },
        behavior: {
          ...resolved.ankiConnect.behavior,
          ...(isObject(ac.behavior)
            ? (ac.behavior as ResolvedConfig['ankiConnect']['behavior'])
            : {}),
        },
        metadata: {
          ...resolved.ankiConnect.metadata,
          ...(isObject(ac.metadata)
            ? (ac.metadata as ResolvedConfig['ankiConnect']['metadata'])
            : {}),
        },
        isLapis: {
          ...resolved.ankiConnect.isLapis,
        },
        isKiku: {
          ...resolved.ankiConnect.isKiku,
          ...(isObject(ac.isKiku) ? (ac.isKiku as ResolvedConfig['ankiConnect']['isKiku']) : {}),
        },
      };

      if (isObject(ac.isLapis)) {
        const lapisEnabled = asBoolean(ac.isLapis.enabled);
        if (lapisEnabled !== undefined) {
          resolved.ankiConnect.isLapis.enabled = lapisEnabled;
        } else if (ac.isLapis.enabled !== undefined) {
          warn(
            'ankiConnect.isLapis.enabled',
            ac.isLapis.enabled,
            resolved.ankiConnect.isLapis.enabled,
            'Expected boolean.',
          );
        }

        const sentenceCardModel = asString(ac.isLapis.sentenceCardModel);
        if (sentenceCardModel !== undefined) {
          resolved.ankiConnect.isLapis.sentenceCardModel = sentenceCardModel;
        } else if (ac.isLapis.sentenceCardModel !== undefined) {
          warn(
            'ankiConnect.isLapis.sentenceCardModel',
            ac.isLapis.sentenceCardModel,
            resolved.ankiConnect.isLapis.sentenceCardModel,
            'Expected string.',
          );
        }

        if (ac.isLapis.sentenceCardSentenceField !== undefined) {
          warn(
            'ankiConnect.isLapis.sentenceCardSentenceField',
            ac.isLapis.sentenceCardSentenceField,
            'Sentence',
            'Deprecated key; sentence-card sentence field is fixed to Sentence.',
          );
        }

        if (ac.isLapis.sentenceCardAudioField !== undefined) {
          warn(
            'ankiConnect.isLapis.sentenceCardAudioField',
            ac.isLapis.sentenceCardAudioField,
            'SentenceAudio',
            'Deprecated key; sentence-card audio field is fixed to SentenceAudio.',
          );
        }
      } else if (ac.isLapis !== undefined) {
        warn('ankiConnect.isLapis', ac.isLapis, resolved.ankiConnect.isLapis, 'Expected object.');
      }

      if (Array.isArray(ac.tags)) {
        const normalizedTags = ac.tags
          .filter((entry): entry is string => typeof entry === 'string')
          .map((entry) => entry.trim())
          .filter((entry) => entry.length > 0);
        if (normalizedTags.length === ac.tags.length) {
          resolved.ankiConnect.tags = [...new Set(normalizedTags)];
        } else {
          resolved.ankiConnect.tags = DEFAULT_CONFIG.ankiConnect.tags;
          warn(
            'ankiConnect.tags',
            ac.tags,
            resolved.ankiConnect.tags,
            'Expected an array of non-empty strings.',
          );
        }
      } else if (ac.tags !== undefined) {
        resolved.ankiConnect.tags = DEFAULT_CONFIG.ankiConnect.tags;
        warn(
          'ankiConnect.tags',
          ac.tags,
          resolved.ankiConnect.tags,
          'Expected an array of strings.',
        );
      }

      const legacy = ac as Record<string, unknown>;
      const hasOwn = (obj: Record<string, unknown>, key: string): boolean =>
        Object.prototype.hasOwnProperty.call(obj, key);
      const asIntegerInRange = (value: unknown, min: number, max: number): number | undefined => {
        const parsed = asNumber(value);
        if (parsed === undefined || !Number.isInteger(parsed) || parsed < min || parsed > max) {
          return undefined;
        }
        return parsed;
      };
      const asPositiveInteger = (value: unknown): number | undefined => {
        const parsed = asNumber(value);
        if (parsed === undefined || !Number.isInteger(parsed) || parsed <= 0) {
          return undefined;
        }
        return parsed;
      };
      const asPositiveNumber = (value: unknown): number | undefined => {
        const parsed = asNumber(value);
        if (parsed === undefined || parsed <= 0) {
          return undefined;
        }
        return parsed;
      };
      const asNonNegativeNumber = (value: unknown): number | undefined => {
        const parsed = asNumber(value);
        if (parsed === undefined || parsed < 0) {
          return undefined;
        }
        return parsed;
      };
      const asImageType = (value: unknown): 'static' | 'avif' | undefined => {
        return value === 'static' || value === 'avif' ? value : undefined;
      };
      const asImageFormat = (value: unknown): 'jpg' | 'png' | 'webp' | undefined => {
        return value === 'jpg' || value === 'png' || value === 'webp' ? value : undefined;
      };
      const asMediaInsertMode = (value: unknown): 'append' | 'prepend' | undefined => {
        return value === 'append' || value === 'prepend' ? value : undefined;
      };
      const asNotificationType = (
        value: unknown,
      ): 'osd' | 'system' | 'both' | 'none' | undefined => {
        return value === 'osd' || value === 'system' || value === 'both' || value === 'none'
          ? value
          : undefined;
      };
      const mapLegacy = <T>(
        key: string,
        parse: (value: unknown) => T | undefined,
        apply: (value: T) => void,
        fallback: unknown,
        message: string,
      ): void => {
        const value = legacy[key];
        if (value === undefined) return;
        const parsed = parse(value);
        if (parsed === undefined) {
          warn(`ankiConnect.${key}`, value, fallback, message);
          return;
        }
        apply(parsed);
      };

      if (!hasOwn(fields, 'audio')) {
        mapLegacy(
          'audioField',
          asString,
          (value) => {
            resolved.ankiConnect.fields.audio = value;
          },
          resolved.ankiConnect.fields.audio,
          'Expected string.',
        );
      }
      if (!hasOwn(fields, 'image')) {
        mapLegacy(
          'imageField',
          asString,
          (value) => {
            resolved.ankiConnect.fields.image = value;
          },
          resolved.ankiConnect.fields.image,
          'Expected string.',
        );
      }
      if (!hasOwn(fields, 'sentence')) {
        mapLegacy(
          'sentenceField',
          asString,
          (value) => {
            resolved.ankiConnect.fields.sentence = value;
          },
          resolved.ankiConnect.fields.sentence,
          'Expected string.',
        );
      }
      if (!hasOwn(fields, 'miscInfo')) {
        mapLegacy(
          'miscInfoField',
          asString,
          (value) => {
            resolved.ankiConnect.fields.miscInfo = value;
          },
          resolved.ankiConnect.fields.miscInfo,
          'Expected string.',
        );
      }
      if (!hasOwn(metadata, 'pattern')) {
        mapLegacy(
          'miscInfoPattern',
          asString,
          (value) => {
            resolved.ankiConnect.metadata.pattern = value;
          },
          resolved.ankiConnect.metadata.pattern,
          'Expected string.',
        );
      }
      if (!hasOwn(media, 'generateAudio')) {
        mapLegacy(
          'generateAudio',
          asBoolean,
          (value) => {
            resolved.ankiConnect.media.generateAudio = value;
          },
          resolved.ankiConnect.media.generateAudio,
          'Expected boolean.',
        );
      }
      if (!hasOwn(media, 'generateImage')) {
        mapLegacy(
          'generateImage',
          asBoolean,
          (value) => {
            resolved.ankiConnect.media.generateImage = value;
          },
          resolved.ankiConnect.media.generateImage,
          'Expected boolean.',
        );
      }
      if (!hasOwn(media, 'imageType')) {
        mapLegacy(
          'imageType',
          asImageType,
          (value) => {
            resolved.ankiConnect.media.imageType = value;
          },
          resolved.ankiConnect.media.imageType,
          "Expected 'static' or 'avif'.",
        );
      }
      if (!hasOwn(media, 'imageFormat')) {
        mapLegacy(
          'imageFormat',
          asImageFormat,
          (value) => {
            resolved.ankiConnect.media.imageFormat = value;
          },
          resolved.ankiConnect.media.imageFormat,
          "Expected 'jpg', 'png', or 'webp'.",
        );
      }
      if (!hasOwn(media, 'imageQuality')) {
        mapLegacy(
          'imageQuality',
          (value) => asIntegerInRange(value, 1, 100),
          (value) => {
            resolved.ankiConnect.media.imageQuality = value;
          },
          resolved.ankiConnect.media.imageQuality,
          'Expected integer between 1 and 100.',
        );
      }
      if (!hasOwn(media, 'imageMaxWidth')) {
        mapLegacy(
          'imageMaxWidth',
          asPositiveInteger,
          (value) => {
            resolved.ankiConnect.media.imageMaxWidth = value;
          },
          resolved.ankiConnect.media.imageMaxWidth,
          'Expected positive integer.',
        );
      }
      if (!hasOwn(media, 'imageMaxHeight')) {
        mapLegacy(
          'imageMaxHeight',
          asPositiveInteger,
          (value) => {
            resolved.ankiConnect.media.imageMaxHeight = value;
          },
          resolved.ankiConnect.media.imageMaxHeight,
          'Expected positive integer.',
        );
      }
      if (!hasOwn(media, 'animatedFps')) {
        mapLegacy(
          'animatedFps',
          (value) => asIntegerInRange(value, 1, 60),
          (value) => {
            resolved.ankiConnect.media.animatedFps = value;
          },
          resolved.ankiConnect.media.animatedFps,
          'Expected integer between 1 and 60.',
        );
      }
      if (!hasOwn(media, 'animatedMaxWidth')) {
        mapLegacy(
          'animatedMaxWidth',
          asPositiveInteger,
          (value) => {
            resolved.ankiConnect.media.animatedMaxWidth = value;
          },
          resolved.ankiConnect.media.animatedMaxWidth,
          'Expected positive integer.',
        );
      }
      if (!hasOwn(media, 'animatedMaxHeight')) {
        mapLegacy(
          'animatedMaxHeight',
          asPositiveInteger,
          (value) => {
            resolved.ankiConnect.media.animatedMaxHeight = value;
          },
          resolved.ankiConnect.media.animatedMaxHeight,
          'Expected positive integer.',
        );
      }
      if (!hasOwn(media, 'animatedCrf')) {
        mapLegacy(
          'animatedCrf',
          (value) => asIntegerInRange(value, 0, 63),
          (value) => {
            resolved.ankiConnect.media.animatedCrf = value;
          },
          resolved.ankiConnect.media.animatedCrf,
          'Expected integer between 0 and 63.',
        );
      }
      if (!hasOwn(media, 'audioPadding')) {
        mapLegacy(
          'audioPadding',
          asNonNegativeNumber,
          (value) => {
            resolved.ankiConnect.media.audioPadding = value;
          },
          resolved.ankiConnect.media.audioPadding,
          'Expected non-negative number.',
        );
      }
      if (!hasOwn(media, 'fallbackDuration')) {
        mapLegacy(
          'fallbackDuration',
          asPositiveNumber,
          (value) => {
            resolved.ankiConnect.media.fallbackDuration = value;
          },
          resolved.ankiConnect.media.fallbackDuration,
          'Expected positive number.',
        );
      }
      if (!hasOwn(media, 'maxMediaDuration')) {
        mapLegacy(
          'maxMediaDuration',
          asNonNegativeNumber,
          (value) => {
            resolved.ankiConnect.media.maxMediaDuration = value;
          },
          resolved.ankiConnect.media.maxMediaDuration,
          'Expected non-negative number.',
        );
      }
      if (!hasOwn(behavior, 'overwriteAudio')) {
        mapLegacy(
          'overwriteAudio',
          asBoolean,
          (value) => {
            resolved.ankiConnect.behavior.overwriteAudio = value;
          },
          resolved.ankiConnect.behavior.overwriteAudio,
          'Expected boolean.',
        );
      }
      if (!hasOwn(behavior, 'overwriteImage')) {
        mapLegacy(
          'overwriteImage',
          asBoolean,
          (value) => {
            resolved.ankiConnect.behavior.overwriteImage = value;
          },
          resolved.ankiConnect.behavior.overwriteImage,
          'Expected boolean.',
        );
      }
      if (!hasOwn(behavior, 'mediaInsertMode')) {
        mapLegacy(
          'mediaInsertMode',
          asMediaInsertMode,
          (value) => {
            resolved.ankiConnect.behavior.mediaInsertMode = value;
          },
          resolved.ankiConnect.behavior.mediaInsertMode,
          "Expected 'append' or 'prepend'.",
        );
      }
      if (!hasOwn(behavior, 'highlightWord')) {
        mapLegacy(
          'highlightWord',
          asBoolean,
          (value) => {
            resolved.ankiConnect.behavior.highlightWord = value;
          },
          resolved.ankiConnect.behavior.highlightWord,
          'Expected boolean.',
        );
      }
      if (!hasOwn(behavior, 'notificationType')) {
        mapLegacy(
          'notificationType',
          asNotificationType,
          (value) => {
            resolved.ankiConnect.behavior.notificationType = value;
          },
          resolved.ankiConnect.behavior.notificationType,
          "Expected 'osd', 'system', 'both', or 'none'.",
        );
      }
      if (!hasOwn(behavior, 'autoUpdateNewCards')) {
        mapLegacy(
          'autoUpdateNewCards',
          asBoolean,
          (value) => {
            resolved.ankiConnect.behavior.autoUpdateNewCards = value;
          },
          resolved.ankiConnect.behavior.autoUpdateNewCards,
          'Expected boolean.',
        );
      }

      const nPlusOneConfig = isObject(ac.nPlusOne) ? (ac.nPlusOne as Record<string, unknown>) : {};

      const nPlusOneHighlightEnabled = asBoolean(nPlusOneConfig.highlightEnabled);
      if (nPlusOneHighlightEnabled !== undefined) {
        resolved.ankiConnect.nPlusOne.highlightEnabled = nPlusOneHighlightEnabled;
      } else {
        const legacyNPlusOneHighlightEnabled = asBoolean(behavior.nPlusOneHighlightEnabled);
        if (legacyNPlusOneHighlightEnabled !== undefined) {
          resolved.ankiConnect.nPlusOne.highlightEnabled = legacyNPlusOneHighlightEnabled;
          warn(
            'ankiConnect.behavior.nPlusOneHighlightEnabled',
            behavior.nPlusOneHighlightEnabled,
            DEFAULT_CONFIG.ankiConnect.nPlusOne.highlightEnabled,
            'Legacy key is deprecated; use ankiConnect.nPlusOne.highlightEnabled',
          );
        } else if (nPlusOneConfig.highlightEnabled !== undefined) {
          warn(
            'ankiConnect.nPlusOne.highlightEnabled',
            nPlusOneConfig.highlightEnabled,
            resolved.ankiConnect.nPlusOne.highlightEnabled,
            'Expected boolean.',
          );
          resolved.ankiConnect.nPlusOne.highlightEnabled =
            DEFAULT_CONFIG.ankiConnect.nPlusOne.highlightEnabled;
        } else {
          resolved.ankiConnect.nPlusOne.highlightEnabled =
            DEFAULT_CONFIG.ankiConnect.nPlusOne.highlightEnabled;
        }
      }

      const nPlusOneRefreshMinutes = asNumber(nPlusOneConfig.refreshMinutes);
      const hasValidNPlusOneRefreshMinutes =
        nPlusOneRefreshMinutes !== undefined &&
        Number.isInteger(nPlusOneRefreshMinutes) &&
        nPlusOneRefreshMinutes > 0;
      if (nPlusOneRefreshMinutes !== undefined) {
        if (hasValidNPlusOneRefreshMinutes) {
          resolved.ankiConnect.nPlusOne.refreshMinutes = nPlusOneRefreshMinutes;
        } else {
          warn(
            'ankiConnect.nPlusOne.refreshMinutes',
            nPlusOneConfig.refreshMinutes,
            resolved.ankiConnect.nPlusOne.refreshMinutes,
            'Expected a positive integer.',
          );
          resolved.ankiConnect.nPlusOne.refreshMinutes =
            DEFAULT_CONFIG.ankiConnect.nPlusOne.refreshMinutes;
        }
      } else if (asNumber(behavior.nPlusOneRefreshMinutes) !== undefined) {
        const legacyNPlusOneRefreshMinutes = asNumber(behavior.nPlusOneRefreshMinutes);
        const hasValidLegacyRefreshMinutes =
          legacyNPlusOneRefreshMinutes !== undefined &&
          Number.isInteger(legacyNPlusOneRefreshMinutes) &&
          legacyNPlusOneRefreshMinutes > 0;
        if (hasValidLegacyRefreshMinutes) {
          resolved.ankiConnect.nPlusOne.refreshMinutes = legacyNPlusOneRefreshMinutes;
          warn(
            'ankiConnect.behavior.nPlusOneRefreshMinutes',
            behavior.nPlusOneRefreshMinutes,
            DEFAULT_CONFIG.ankiConnect.nPlusOne.refreshMinutes,
            'Legacy key is deprecated; use ankiConnect.nPlusOne.refreshMinutes',
          );
        } else {
          warn(
            'ankiConnect.behavior.nPlusOneRefreshMinutes',
            behavior.nPlusOneRefreshMinutes,
            resolved.ankiConnect.nPlusOne.refreshMinutes,
            'Expected a positive integer.',
          );
          resolved.ankiConnect.nPlusOne.refreshMinutes =
            DEFAULT_CONFIG.ankiConnect.nPlusOne.refreshMinutes;
        }
      } else {
        resolved.ankiConnect.nPlusOne.refreshMinutes =
          DEFAULT_CONFIG.ankiConnect.nPlusOne.refreshMinutes;
      }

      const nPlusOneMinSentenceWords = asNumber(nPlusOneConfig.minSentenceWords);
      const hasValidNPlusOneMinSentenceWords =
        nPlusOneMinSentenceWords !== undefined &&
        Number.isInteger(nPlusOneMinSentenceWords) &&
        nPlusOneMinSentenceWords > 0;
      if (nPlusOneMinSentenceWords !== undefined) {
        if (hasValidNPlusOneMinSentenceWords) {
          resolved.ankiConnect.nPlusOne.minSentenceWords = nPlusOneMinSentenceWords;
        } else {
          warn(
            'ankiConnect.nPlusOne.minSentenceWords',
            nPlusOneConfig.minSentenceWords,
            resolved.ankiConnect.nPlusOne.minSentenceWords,
            'Expected a positive integer.',
          );
          resolved.ankiConnect.nPlusOne.minSentenceWords =
            DEFAULT_CONFIG.ankiConnect.nPlusOne.minSentenceWords;
        }
      } else {
        resolved.ankiConnect.nPlusOne.minSentenceWords =
          DEFAULT_CONFIG.ankiConnect.nPlusOne.minSentenceWords;
      }

      const nPlusOneMatchMode = asString(nPlusOneConfig.matchMode);
      const legacyNPlusOneMatchMode = asString(behavior.nPlusOneMatchMode);
      const hasValidNPlusOneMatchMode =
        nPlusOneMatchMode === 'headword' || nPlusOneMatchMode === 'surface';
      const hasValidLegacyMatchMode =
        legacyNPlusOneMatchMode === 'headword' || legacyNPlusOneMatchMode === 'surface';
      if (hasValidNPlusOneMatchMode) {
        resolved.ankiConnect.nPlusOne.matchMode = nPlusOneMatchMode;
      } else if (nPlusOneMatchMode !== undefined) {
        warn(
          'ankiConnect.nPlusOne.matchMode',
          nPlusOneConfig.matchMode,
          DEFAULT_CONFIG.ankiConnect.nPlusOne.matchMode,
          "Expected 'headword' or 'surface'.",
        );
        resolved.ankiConnect.nPlusOne.matchMode = DEFAULT_CONFIG.ankiConnect.nPlusOne.matchMode;
      } else if (legacyNPlusOneMatchMode !== undefined) {
        if (hasValidLegacyMatchMode) {
          resolved.ankiConnect.nPlusOne.matchMode = legacyNPlusOneMatchMode;
          warn(
            'ankiConnect.behavior.nPlusOneMatchMode',
            behavior.nPlusOneMatchMode,
            DEFAULT_CONFIG.ankiConnect.nPlusOne.matchMode,
            'Legacy key is deprecated; use ankiConnect.nPlusOne.matchMode',
          );
        } else {
          warn(
            'ankiConnect.behavior.nPlusOneMatchMode',
            behavior.nPlusOneMatchMode,
            resolved.ankiConnect.nPlusOne.matchMode,
            "Expected 'headword' or 'surface'.",
          );
          resolved.ankiConnect.nPlusOne.matchMode = DEFAULT_CONFIG.ankiConnect.nPlusOne.matchMode;
        }
      } else {
        resolved.ankiConnect.nPlusOne.matchMode = DEFAULT_CONFIG.ankiConnect.nPlusOne.matchMode;
      }

      const nPlusOneDecks = nPlusOneConfig.decks;
      if (Array.isArray(nPlusOneDecks)) {
        const normalizedDecks = nPlusOneDecks
          .filter((entry): entry is string => typeof entry === 'string')
          .map((entry) => entry.trim())
          .filter((entry) => entry.length > 0);

        if (normalizedDecks.length === nPlusOneDecks.length) {
          resolved.ankiConnect.nPlusOne.decks = [...new Set(normalizedDecks)];
        } else if (nPlusOneDecks.length > 0) {
          warn(
            'ankiConnect.nPlusOne.decks',
            nPlusOneDecks,
            resolved.ankiConnect.nPlusOne.decks,
            'Expected an array of strings.',
          );
        } else {
          resolved.ankiConnect.nPlusOne.decks = [];
        }
      } else if (nPlusOneDecks !== undefined) {
        warn(
          'ankiConnect.nPlusOne.decks',
          nPlusOneDecks,
          resolved.ankiConnect.nPlusOne.decks,
          'Expected an array of strings.',
        );
        resolved.ankiConnect.nPlusOne.decks = [];
      }

      const nPlusOneHighlightColor = asColor(nPlusOneConfig.nPlusOne);
      if (nPlusOneHighlightColor !== undefined) {
        resolved.ankiConnect.nPlusOne.nPlusOne = nPlusOneHighlightColor;
      } else if (nPlusOneConfig.nPlusOne !== undefined) {
        warn(
          'ankiConnect.nPlusOne.nPlusOne',
          nPlusOneConfig.nPlusOne,
          resolved.ankiConnect.nPlusOne.nPlusOne,
          'Expected a hex color value.',
        );
        resolved.ankiConnect.nPlusOne.nPlusOne = DEFAULT_CONFIG.ankiConnect.nPlusOne.nPlusOne;
      }

      const nPlusOneKnownWordColor = asColor(nPlusOneConfig.knownWord);
      if (nPlusOneKnownWordColor !== undefined) {
        resolved.ankiConnect.nPlusOne.knownWord = nPlusOneKnownWordColor;
      } else if (nPlusOneConfig.knownWord !== undefined) {
        warn(
          'ankiConnect.nPlusOne.knownWord',
          nPlusOneConfig.knownWord,
          resolved.ankiConnect.nPlusOne.knownWord,
          'Expected a hex color value.',
        );
        resolved.ankiConnect.nPlusOne.knownWord = DEFAULT_CONFIG.ankiConnect.nPlusOne.knownWord;
      }

      if (
        resolved.ankiConnect.isKiku.fieldGrouping !== 'auto' &&
        resolved.ankiConnect.isKiku.fieldGrouping !== 'manual' &&
        resolved.ankiConnect.isKiku.fieldGrouping !== 'disabled'
      ) {
        warn(
          'ankiConnect.isKiku.fieldGrouping',
          resolved.ankiConnect.isKiku.fieldGrouping,
          DEFAULT_CONFIG.ankiConnect.isKiku.fieldGrouping,
          'Expected auto, manual, or disabled.',
        );
        resolved.ankiConnect.isKiku.fieldGrouping = DEFAULT_CONFIG.ankiConnect.isKiku.fieldGrouping;
      }
    }

    return { resolved, warnings };
  }
}
