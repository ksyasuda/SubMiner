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

export function parsePluginRuntimeConfigContent(content: string): PluginRuntimeConfig {
  const runtimeConfig: PluginRuntimeConfig = { socketPath: DEFAULT_SOCKET_PATH };
  for (const line of content.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (trimmed.length === 0 || trimmed.startsWith('#')) continue;
    const socketMatch = trimmed.match(/^socket_path\s*=\s*(.+)$/i);
    if (!socketMatch) continue;
    const value = (socketMatch[1] || '').split('#', 1)[0]?.trim() || '';
    if (value) runtimeConfig.socketPath = value;
  }
  return runtimeConfig;
}

export function readPluginRuntimeConfig(logLevel: LogLevel): PluginRuntimeConfig {
  const candidates = getPluginConfigCandidates();
  const defaults: PluginRuntimeConfig = { socketPath: DEFAULT_SOCKET_PATH };

  for (const configPath of candidates) {
    if (!fs.existsSync(configPath)) continue;
    try {
      const parsed = parsePluginRuntimeConfigContent(fs.readFileSync(configPath, 'utf8'));
      log(
        'debug',
        logLevel,
        `Using mpv plugin settings from ${configPath}: socket_path=${parsed.socketPath}`,
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
    `No mpv subminer.conf found; using launcher defaults (socket_path=${defaults.socketPath})`,
  );
  return defaults;
}
