import { log } from '../log.js';
import type { Backend, LogLevel, PluginRuntimeConfig } from '../types.js';
import { DEFAULT_SOCKET_PATH } from '../types.js';
import { buildSubminerPluginRuntimeScriptOptParts } from '../../src/shared/subminer-plugin-script-opts.js';
import { parseLauncherMpvConfig } from './mpv-config.js';
import { readLauncherMainConfigObject } from './shared-config-reader.js';

function rootObject(root: Record<string, unknown> | null, key: string): Record<string, unknown> {
  const value = root?.[key];
  return value && typeof value === 'object' && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : {};
}

function booleanOrDefault(value: unknown, fallback: boolean): boolean {
  return typeof value === 'boolean' ? value : fallback;
}

function nonEmptyStringOrDefault(value: unknown, fallback: string): string {
  if (typeof value !== 'string') return fallback;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : fallback;
}

function validBackendOrDefault(value: unknown, fallback: Backend): Backend {
  if (typeof value !== 'string') return fallback;
  const normalized = value.trim().toLowerCase();
  if (
    normalized === 'auto' ||
    normalized === 'hyprland' ||
    normalized === 'sway' ||
    normalized === 'x11' ||
    normalized === 'macos' ||
    normalized === 'windows'
  ) {
    return normalized;
  }
  return fallback;
}

export function parsePluginRuntimeConfigFromMainConfig(
  root: Record<string, unknown> | null,
  logLevel: LogLevel = 'info',
): PluginRuntimeConfig {
  const mpvConfig = root ? parseLauncherMpvConfig(root) : {};
  const texthooker = rootObject(root, 'texthooker');

  return {
    socketPath: mpvConfig.socketPath ?? DEFAULT_SOCKET_PATH,
    binaryPath: mpvConfig.subminerBinaryPath ?? '',
    backend: validBackendOrDefault(mpvConfig.backend, 'auto'),
    logLevel,
    autoStart: booleanOrDefault(mpvConfig.autoStartSubMiner, true),
    autoStartVisibleOverlay: booleanOrDefault(root?.auto_start_overlay, false),
    autoStartPauseUntilReady: booleanOrDefault(mpvConfig.pauseUntilOverlayReady, true),
    texthookerEnabled: booleanOrDefault(texthooker.launchAtStartup, false),
  };
}

export function buildPluginRuntimeScriptOptParts(
  runtimeConfig: PluginRuntimeConfig,
  fallbackAppPath: string,
): string[] {
  return buildSubminerPluginRuntimeScriptOptParts(runtimeConfig, fallbackAppPath);
}

export function readPluginRuntimeConfig(logLevel: LogLevel): PluginRuntimeConfig {
  const parsed = parsePluginRuntimeConfigFromMainConfig(readLauncherMainConfigObject(), logLevel);

  log(
    'debug',
    logLevel,
    `Using mpv plugin settings from SubMiner config: socket_path=${parsed.socketPath}, backend=${parsed.backend}, auto_start=${parsed.autoStart}, auto_start_visible_overlay=${parsed.autoStartVisibleOverlay}, auto_start_pause_until_ready=${parsed.autoStartPauseUntilReady}, texthooker_enabled=${parsed.texthookerEnabled}`,
  );
  return parsed;
}
