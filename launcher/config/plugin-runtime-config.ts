import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { log } from '../log.js';
import type { LogLevel, PluginRuntimeConfig } from '../types.js';
import { DEFAULT_SOCKET_PATH } from '../types.js';

export function getPluginConfigCandidates(): string[] {
  const xdgConfigHome = process.env.XDG_CONFIG_HOME || path.join(os.homedir(), '.config');
  return Array.from(
    new Set([
      path.join(xdgConfigHome, 'mpv', 'script-opts', 'subminer.conf'),
      path.join(os.homedir(), '.config', 'mpv', 'script-opts', 'subminer.conf'),
    ]),
  );
}

export function parsePluginRuntimeConfigContent(
  content: string,
  logLevel: LogLevel = 'warn',
): PluginRuntimeConfig {
  const runtimeConfig: PluginRuntimeConfig = {
    socketPath: DEFAULT_SOCKET_PATH,
    autoStart: true,
    autoStartVisibleOverlay: true,
    autoStartPauseUntilReady: true,
  };

  const parseBooleanValue = (key: string, value: string): boolean => {
    const normalized = value.trim().toLowerCase();
    if (['yes', 'true', '1', 'on'].includes(normalized)) return true;
    if (['no', 'false', '0', 'off'].includes(normalized)) return false;
    log('warn', logLevel, `Invalid boolean value for ${key}: "${value}". Using false.`);
    return false;
  };

  for (const line of content.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (trimmed.length === 0 || trimmed.startsWith('#')) continue;
    const keyValueMatch = trimmed.match(/^([a-z0-9_-]+)\s*=\s*(.+)$/i);
    if (!keyValueMatch) continue;
    const key = (keyValueMatch[1] || '').toLowerCase();
    const value = (keyValueMatch[2] || '').split('#', 1)[0]?.trim() || '';
    if (!value) continue;

    if (key === 'socket_path') {
      runtimeConfig.socketPath = value;
      continue;
    }
    if (key === 'auto_start') {
      runtimeConfig.autoStart = parseBooleanValue('auto_start', value);
      continue;
    }
    if (key === 'auto_start_visible_overlay') {
      runtimeConfig.autoStartVisibleOverlay = parseBooleanValue(
        'auto_start_visible_overlay',
        value,
      );
      continue;
    }
    if (key === 'auto_start_pause_until_ready') {
      runtimeConfig.autoStartPauseUntilReady = parseBooleanValue(
        'auto_start_pause_until_ready',
        value,
      );
    }
  }
  return runtimeConfig;
}

export function readPluginRuntimeConfig(logLevel: LogLevel): PluginRuntimeConfig {
  const candidates = getPluginConfigCandidates();
  const defaults: PluginRuntimeConfig = {
    socketPath: DEFAULT_SOCKET_PATH,
    autoStart: true,
    autoStartVisibleOverlay: true,
    autoStartPauseUntilReady: true,
  };

  for (const configPath of candidates) {
    if (!fs.existsSync(configPath)) continue;
    try {
      const parsed = parsePluginRuntimeConfigContent(fs.readFileSync(configPath, 'utf8'));
      log(
        'debug',
        logLevel,
        `Using mpv plugin settings from ${configPath}: socket_path=${parsed.socketPath}, auto_start=${parsed.autoStart}, auto_start_visible_overlay=${parsed.autoStartVisibleOverlay}, auto_start_pause_until_ready=${parsed.autoStartPauseUntilReady}`,
      );
      return parsed;
    } catch {
      log('warn', logLevel, `Failed to read ${configPath}; using launcher defaults`);
      return defaults;
    }
  }

  log(
    'debug',
    logLevel,
    `No mpv subminer.conf found; using launcher defaults (socket_path=${defaults.socketPath}, auto_start=${defaults.autoStart}, auto_start_visible_overlay=${defaults.autoStartVisibleOverlay}, auto_start_pause_until_ready=${defaults.autoStartPauseUntilReady})`,
  );
  return defaults;
}
