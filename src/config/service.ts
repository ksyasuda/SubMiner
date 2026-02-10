import * as fs from "fs";
import * as path from "path";
import { parse as parseJsonc } from "jsonc-parser";
import {
  Config,
  ConfigValidationWarning,
  RawConfig,
  ResolvedConfig,
} from "../types";
import {
  DEFAULT_CONFIG,
  deepCloneConfig,
  deepMergeRawConfig,
} from "./definitions";

interface LoadResult {
  config: RawConfig;
  path: string;
}

function isObject(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function asNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value)
    ? value
    : undefined;
}

function asString(value: unknown): string | undefined {
  return typeof value === "string" ? value : undefined;
}

function asBoolean(value: unknown): boolean | undefined {
  return typeof value === "boolean" ? value : undefined;
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
    this.configFileJsonc = path.join(configDir, "config.jsonc");
    this.configFileJson = path.join(configDir, "config.json");
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

  saveRawConfig(config: RawConfig): void {
    if (!fs.existsSync(this.configDir)) {
      fs.mkdirSync(this.configDir, { recursive: true });
    }
    const targetPath = this.configPathInUse.endsWith(".json")
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
    const configPath = fs.existsSync(this.configFileJsonc)
      ? this.configFileJsonc
      : fs.existsSync(this.configFileJson)
        ? this.configFileJson
        : this.configFileJsonc;

    if (!fs.existsSync(configPath)) {
      return { config: {}, path: configPath };
    }

    try {
      const data = fs.readFileSync(configPath, "utf-8");
      const parsed = configPath.endsWith(".jsonc")
        ? parseJsonc(data)
        : JSON.parse(data);
      return {
        config: isObject(parsed) ? (parsed as Config) : {},
        path: configPath,
      };
    } catch {
      return { config: {}, path: configPath };
    }
  }

  private resolveConfig(raw: RawConfig): {
    resolved: ResolvedConfig;
    warnings: ConfigValidationWarning[];
  } {
    const warnings: ConfigValidationWarning[] = [];
    const resolved = deepCloneConfig(DEFAULT_CONFIG);

    const warn = (
      path: string,
      value: unknown,
      fallback: unknown,
      message: string,
    ): void => {
      warnings.push({
        path,
        value,
        fallback,
        message,
      });
    };

    const src = isObject(raw) ? raw : {};

    if (isObject(src.texthooker)) {
      const openBrowser = asBoolean(src.texthooker.openBrowser);
      if (openBrowser !== undefined) {
        resolved.texthooker.openBrowser = openBrowser;
      } else if (src.texthooker.openBrowser !== undefined) {
        warn(
          "texthooker.openBrowser",
          src.texthooker.openBrowser,
          resolved.texthooker.openBrowser,
          "Expected boolean.",
        );
      }
    }

    if (isObject(src.websocket)) {
      const enabled = src.websocket.enabled;
      if (enabled === "auto" || enabled === true || enabled === false) {
        resolved.websocket.enabled = enabled;
      } else if (enabled !== undefined) {
        warn(
          "websocket.enabled",
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
          "websocket.port",
          src.websocket.port,
          resolved.websocket.port,
          "Expected integer between 1 and 65535.",
        );
      }
    }

    if (Array.isArray(src.keybindings)) {
      resolved.keybindings = src.keybindings.filter(
        (
          entry,
        ): entry is { key: string; command: (string | number)[] | null } => {
          if (!isObject(entry)) return false;
          if (typeof entry.key !== "string") return false;
          if (entry.command === null) return true;
          return Array.isArray(entry.command);
        },
      );
    }

    if (isObject(src.shortcuts)) {
      const shortcutKeys = [
        "toggleVisibleOverlayGlobal",
        "toggleInvisibleOverlayGlobal",
        "copySubtitle",
        "copySubtitleMultiple",
        "updateLastCardFromClipboard",
        "triggerFieldGrouping",
        "triggerSubsync",
        "mineSentence",
        "mineSentenceMultiple",
        "toggleSecondarySub",
        "markAudioCard",
        "openRuntimeOptions",
      ] as const;

      for (const key of shortcutKeys) {
        const value = src.shortcuts[key];
        if (typeof value === "string" || value === null) {
          resolved.shortcuts[key] =
            value as (typeof resolved.shortcuts)[typeof key];
        } else if (value !== undefined) {
          warn(
            `shortcuts.${key}`,
            value,
            resolved.shortcuts[key],
            "Expected string or null.",
          );
        }
      }

      const timeout = asNumber(src.shortcuts.multiCopyTimeoutMs);
      if (timeout !== undefined && timeout > 0) {
        resolved.shortcuts.multiCopyTimeoutMs = Math.floor(timeout);
      } else if (src.shortcuts.multiCopyTimeoutMs !== undefined) {
        warn(
          "shortcuts.multiCopyTimeoutMs",
          src.shortcuts.multiCopyTimeoutMs,
          resolved.shortcuts.multiCopyTimeoutMs,
          "Expected positive number.",
        );
      }
    }

    if (isObject(src.invisibleOverlay)) {
      const startupVisibility = src.invisibleOverlay.startupVisibility;
      if (
        startupVisibility === "platform-default" ||
        startupVisibility === "visible" ||
        startupVisibility === "hidden"
      ) {
        resolved.invisibleOverlay.startupVisibility = startupVisibility;
      } else if (startupVisibility !== undefined) {
        warn(
          "invisibleOverlay.startupVisibility",
          startupVisibility,
          resolved.invisibleOverlay.startupVisibility,
          "Expected platform-default, visible, or hidden.",
        );
      }
    }

    if (isObject(src.secondarySub)) {
      if (Array.isArray(src.secondarySub.secondarySubLanguages)) {
        resolved.secondarySub.secondarySubLanguages =
          src.secondarySub.secondarySubLanguages.filter(
            (item): item is string => typeof item === "string",
          );
      }
      const autoLoad = asBoolean(src.secondarySub.autoLoadSecondarySub);
      if (autoLoad !== undefined) {
        resolved.secondarySub.autoLoadSecondarySub = autoLoad;
      }
      const defaultMode = src.secondarySub.defaultMode;
      if (
        defaultMode === "hidden" ||
        defaultMode === "visible" ||
        defaultMode === "hover"
      ) {
        resolved.secondarySub.defaultMode = defaultMode;
      } else if (defaultMode !== undefined) {
        warn(
          "secondarySub.defaultMode",
          defaultMode,
          resolved.secondarySub.defaultMode,
          "Expected hidden, visible, or hover.",
        );
      }
    }

    if (isObject(src.subsync)) {
      const mode = src.subsync.defaultMode;
      if (mode === "auto" || mode === "manual") {
        resolved.subsync.defaultMode = mode;
      } else if (mode !== undefined) {
        warn(
          "subsync.defaultMode",
          mode,
          resolved.subsync.defaultMode,
          "Expected auto or manual.",
        );
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
      if (apiKeyCommand !== undefined)
        resolved.jimaku.apiKeyCommand = apiKeyCommand;
      const apiBaseUrl = asString(src.jimaku.apiBaseUrl);
      if (apiBaseUrl !== undefined) resolved.jimaku.apiBaseUrl = apiBaseUrl;

      const lang = src.jimaku.languagePreference;
      if (lang === "ja" || lang === "en" || lang === "none") {
        resolved.jimaku.languagePreference = lang;
      } else if (lang !== undefined) {
        warn(
          "jimaku.languagePreference",
          lang,
          resolved.jimaku.languagePreference,
          "Expected ja, en, or none.",
        );
      }

      const maxEntryResults = asNumber(src.jimaku.maxEntryResults);
      if (maxEntryResults !== undefined && maxEntryResults > 0) {
        resolved.jimaku.maxEntryResults = Math.floor(maxEntryResults);
      } else if (src.jimaku.maxEntryResults !== undefined) {
        warn(
          "jimaku.maxEntryResults",
          src.jimaku.maxEntryResults,
          resolved.jimaku.maxEntryResults,
          "Expected positive number.",
        );
      }
    }

    if (isObject(src.youtubeSubgen)) {
      const mode = src.youtubeSubgen.mode;
      if (mode === "automatic" || mode === "preprocess" || mode === "off") {
        resolved.youtubeSubgen.mode = mode;
      } else if (mode !== undefined) {
        warn(
          "youtubeSubgen.mode",
          mode,
          resolved.youtubeSubgen.mode,
          "Expected automatic, preprocess, or off.",
        );
      }

      const whisperBin = asString(src.youtubeSubgen.whisperBin);
      if (whisperBin !== undefined) {
        resolved.youtubeSubgen.whisperBin = whisperBin;
      } else if (src.youtubeSubgen.whisperBin !== undefined) {
        warn(
          "youtubeSubgen.whisperBin",
          src.youtubeSubgen.whisperBin,
          resolved.youtubeSubgen.whisperBin,
          "Expected string.",
        );
      }

      const whisperModel = asString(src.youtubeSubgen.whisperModel);
      if (whisperModel !== undefined) {
        resolved.youtubeSubgen.whisperModel = whisperModel;
      } else if (src.youtubeSubgen.whisperModel !== undefined) {
        warn(
          "youtubeSubgen.whisperModel",
          src.youtubeSubgen.whisperModel,
          resolved.youtubeSubgen.whisperModel,
          "Expected string.",
        );
      }

      if (Array.isArray(src.youtubeSubgen.primarySubLanguages)) {
        resolved.youtubeSubgen.primarySubLanguages =
          src.youtubeSubgen.primarySubLanguages.filter(
            (item): item is string => typeof item === "string",
          );
      } else if (src.youtubeSubgen.primarySubLanguages !== undefined) {
        warn(
          "youtubeSubgen.primarySubLanguages",
          src.youtubeSubgen.primarySubLanguages,
          resolved.youtubeSubgen.primarySubLanguages,
          "Expected string array.",
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
        "bind_visible_overlay_to_mpv_sub_visibility",
        src.bind_visible_overlay_to_mpv_sub_visibility,
        resolved.bind_visible_overlay_to_mpv_sub_visibility,
        "Expected boolean.",
      );
    }

    if (isObject(src.subtitleStyle)) {
      resolved.subtitleStyle = {
        ...resolved.subtitleStyle,
        ...(src.subtitleStyle as ResolvedConfig["subtitleStyle"]),
        secondary: {
          ...resolved.subtitleStyle.secondary,
          ...(isObject(src.subtitleStyle.secondary)
            ? (src.subtitleStyle
                .secondary as ResolvedConfig["subtitleStyle"]["secondary"])
            : {}),
        },
      };
    }

    if (isObject(src.ankiConnect)) {
      const ac = src.ankiConnect;
      const aiSource = isObject(ac.ai)
        ? ac.ai
        : isObject(ac.openRouter)
          ? ac.openRouter
          : {};
      resolved.ankiConnect = {
        ...resolved.ankiConnect,
        ...(isObject(ac) ? (ac as Partial<ResolvedConfig["ankiConnect"]>) : {}),
        fields: {
          ...resolved.ankiConnect.fields,
          ...(isObject(ac.fields)
            ? (ac.fields as ResolvedConfig["ankiConnect"]["fields"])
            : {}),
        },
        ai: {
          ...resolved.ankiConnect.ai,
          ...(aiSource as ResolvedConfig["ankiConnect"]["ai"]),
        },
        media: {
          ...resolved.ankiConnect.media,
          ...(isObject(ac.media)
            ? (ac.media as ResolvedConfig["ankiConnect"]["media"])
            : {}),
        },
        behavior: {
          ...resolved.ankiConnect.behavior,
          ...(isObject(ac.behavior)
            ? (ac.behavior as ResolvedConfig["ankiConnect"]["behavior"])
            : {}),
        },
        metadata: {
          ...resolved.ankiConnect.metadata,
          ...(isObject(ac.metadata)
            ? (ac.metadata as ResolvedConfig["ankiConnect"]["metadata"])
            : {}),
        },
        isLapis: {
          ...resolved.ankiConnect.isLapis,
          ...(isObject(ac.isLapis)
            ? (ac.isLapis as ResolvedConfig["ankiConnect"]["isLapis"])
            : {}),
        },
        isKiku: {
          ...resolved.ankiConnect.isKiku,
          ...(isObject(ac.isKiku)
            ? (ac.isKiku as ResolvedConfig["ankiConnect"]["isKiku"])
            : {}),
        },
      };

      const legacy = ac as Record<string, unknown>;
      const mapLegacy = (
        key: string,
        apply: (value: unknown) => void,
      ): void => {
        if (legacy[key] !== undefined) apply(legacy[key]);
      };

      mapLegacy("audioField", (value) => {
        resolved.ankiConnect.fields.audio = value as string;
      });
      mapLegacy("imageField", (value) => {
        resolved.ankiConnect.fields.image = value as string;
      });
      mapLegacy("sentenceField", (value) => {
        resolved.ankiConnect.fields.sentence = value as string;
      });
      mapLegacy("miscInfoField", (value) => {
        resolved.ankiConnect.fields.miscInfo = value as string;
      });
      mapLegacy("miscInfoPattern", (value) => {
        resolved.ankiConnect.metadata.pattern = value as string;
      });
      mapLegacy("generateAudio", (value) => {
        resolved.ankiConnect.media.generateAudio = value as boolean;
      });
      mapLegacy("generateImage", (value) => {
        resolved.ankiConnect.media.generateImage = value as boolean;
      });
      mapLegacy("imageType", (value) => {
        resolved.ankiConnect.media.imageType = value as "static" | "avif";
      });
      mapLegacy("imageFormat", (value) => {
        resolved.ankiConnect.media.imageFormat = value as
          | "jpg"
          | "png"
          | "webp";
      });
      mapLegacy("imageQuality", (value) => {
        resolved.ankiConnect.media.imageQuality = value as number;
      });
      mapLegacy("imageMaxWidth", (value) => {
        resolved.ankiConnect.media.imageMaxWidth = value as number;
      });
      mapLegacy("imageMaxHeight", (value) => {
        resolved.ankiConnect.media.imageMaxHeight = value as number;
      });
      mapLegacy("animatedFps", (value) => {
        resolved.ankiConnect.media.animatedFps = value as number;
      });
      mapLegacy("animatedMaxWidth", (value) => {
        resolved.ankiConnect.media.animatedMaxWidth = value as number;
      });
      mapLegacy("animatedMaxHeight", (value) => {
        resolved.ankiConnect.media.animatedMaxHeight = value as number;
      });
      mapLegacy("animatedCrf", (value) => {
        resolved.ankiConnect.media.animatedCrf = value as number;
      });
      mapLegacy("audioPadding", (value) => {
        resolved.ankiConnect.media.audioPadding = value as number;
      });
      mapLegacy("fallbackDuration", (value) => {
        resolved.ankiConnect.media.fallbackDuration = value as number;
      });
      mapLegacy("maxMediaDuration", (value) => {
        resolved.ankiConnect.media.maxMediaDuration = value as number;
      });
      mapLegacy("overwriteAudio", (value) => {
        resolved.ankiConnect.behavior.overwriteAudio = value as boolean;
      });
      mapLegacy("overwriteImage", (value) => {
        resolved.ankiConnect.behavior.overwriteImage = value as boolean;
      });
      mapLegacy("mediaInsertMode", (value) => {
        resolved.ankiConnect.behavior.mediaInsertMode = value as
          | "append"
          | "prepend";
      });
      mapLegacy("highlightWord", (value) => {
        resolved.ankiConnect.behavior.highlightWord = value as boolean;
      });
      mapLegacy("notificationType", (value) => {
        resolved.ankiConnect.behavior.notificationType = value as
          | "osd"
          | "system"
          | "both"
          | "none";
      });
      mapLegacy("autoUpdateNewCards", (value) => {
        resolved.ankiConnect.behavior.autoUpdateNewCards = value as boolean;
      });

      if (
        resolved.ankiConnect.isKiku.fieldGrouping !== "auto" &&
        resolved.ankiConnect.isKiku.fieldGrouping !== "manual" &&
        resolved.ankiConnect.isKiku.fieldGrouping !== "disabled"
      ) {
        warn(
          "ankiConnect.isKiku.fieldGrouping",
          resolved.ankiConnect.isKiku.fieldGrouping,
          DEFAULT_CONFIG.ankiConnect.isKiku.fieldGrouping,
          "Expected auto, manual, or disabled.",
        );
        resolved.ankiConnect.isKiku.fieldGrouping =
          DEFAULT_CONFIG.ankiConnect.isKiku.fieldGrouping;
      }
    }

    return { resolved, warnings };
  }
}
