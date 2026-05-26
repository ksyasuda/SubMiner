import { fail } from './log.js';
import type {
  Args,
  LauncherLoggingConfig,
  LauncherJellyfinConfig,
  LauncherMpvConfig,
  LauncherYoutubeSubgenConfig,
  LogLevel,
  PluginRuntimeConfig,
} from './types.js';
import { normalizeLogRotation } from '../src/shared/log-files.js';
import {
  applyInvocationsToArgs,
  applyRootOptionsToArgs,
  createDefaultArgs,
} from './config/args-normalizer.js';
import { parseCliPrograms, resolveTopLevelCommand } from './config/cli-parser-builder.js';
import { parseLauncherJellyfinConfig } from './config/jellyfin-config.js';
import { parseLauncherMpvConfig } from './config/mpv-config.js';
import { readPluginRuntimeConfig as readPluginRuntimeConfigValue } from './config/plugin-runtime-config.js';
import { readLauncherMainConfigObject } from './config/shared-config-reader.js';
import { parseLauncherYoutubeSubgenConfig } from './config/youtube-subgen-config.js';

export function readExternalYomitanProfilePath(
  root: Record<string, unknown> | null,
): string | null {
  const yomitan =
    root?.yomitan && typeof root.yomitan === 'object' && !Array.isArray(root.yomitan)
      ? (root.yomitan as Record<string, unknown>)
      : null;
  const externalProfilePath = yomitan?.externalProfilePath;
  if (typeof externalProfilePath !== 'string') {
    return null;
  }
  const trimmed = externalProfilePath.trim();
  return trimmed.length > 0 ? trimmed : null;
}

export function loadLauncherYoutubeSubgenConfig(): LauncherYoutubeSubgenConfig {
  const root = readLauncherMainConfigObject();
  if (!root) return {};
  return parseLauncherYoutubeSubgenConfig(root);
}

export function loadLauncherJellyfinConfig(): LauncherJellyfinConfig {
  const root = readLauncherMainConfigObject();
  if (!root) return {};
  return parseLauncherJellyfinConfig(root);
}

export function loadLauncherMpvConfig(): LauncherMpvConfig {
  const root = readLauncherMainConfigObject();
  if (!root) return {};
  return parseLauncherMpvConfig(root);
}

function parseLogLevelConfig(value: unknown): LogLevel | undefined {
  if (typeof value !== 'string') return undefined;
  const normalized = value.trim().toLowerCase();
  if (
    normalized === 'debug' ||
    normalized === 'info' ||
    normalized === 'warn' ||
    normalized === 'error'
  ) {
    return normalized;
  }
  return undefined;
}

function parseLogRotationConfig(value: unknown): LauncherLoggingConfig['rotation'] {
  return normalizeLogRotation(value);
}

function parseLogFileConfig(value: unknown): boolean | undefined {
  return typeof value === 'boolean' ? value : undefined;
}

export function loadLauncherLoggingConfig(): LauncherLoggingConfig {
  const root = readLauncherMainConfigObject();
  if (!root) return {};
  const logging =
    root.logging && typeof root.logging === 'object' && !Array.isArray(root.logging)
      ? (root.logging as Record<string, unknown>)
      : null;
  const files =
    logging?.files && typeof logging.files === 'object' && !Array.isArray(logging.files)
      ? (logging.files as Record<string, unknown>)
      : null;
  return {
    level: parseLogLevelConfig(logging?.level),
    rotation: parseLogRotationConfig(logging?.rotation),
    files: files
      ? {
          app: parseLogFileConfig(files.app),
          launcher: parseLogFileConfig(files.launcher),
          mpv: parseLogFileConfig(files.mpv),
        }
      : undefined,
  };
}

export function hasLauncherExternalYomitanProfileConfig(): boolean {
  return readExternalYomitanProfilePath(readLauncherMainConfigObject()) !== null;
}

export function readPluginRuntimeConfig(logLevel: LogLevel): PluginRuntimeConfig {
  return readPluginRuntimeConfigValue(logLevel);
}

export function parseArgs(
  argv: string[],
  scriptName: string,
  launcherConfig: LauncherYoutubeSubgenConfig,
  launcherMpvConfig: LauncherMpvConfig = {},
  launcherLoggingConfig: LauncherLoggingConfig = {},
): Args {
  const topLevelCommand = resolveTopLevelCommand(argv);
  const parsed = createDefaultArgs(launcherConfig, launcherMpvConfig, launcherLoggingConfig);

  if (topLevelCommand && (topLevelCommand.name === 'app' || topLevelCommand.name === 'bin')) {
    parsed.appPassthrough = true;
    parsed.appArgs = argv.slice(topLevelCommand.index + 1);
    return parsed;
  }

  let cliResult: ReturnType<typeof parseCliPrograms>;
  try {
    cliResult = parseCliPrograms(argv, scriptName);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    fail(message);
  }

  applyRootOptionsToArgs(parsed, cliResult.options, cliResult.rootTarget);
  applyInvocationsToArgs(parsed, cliResult.invocations);
  return parsed;
}
