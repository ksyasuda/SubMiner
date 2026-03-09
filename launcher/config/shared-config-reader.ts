import fs from 'node:fs';
import os from 'node:os';
import { parse as parseJsonc } from 'jsonc-parser';
import { resolveConfigFilePath } from '../../src/config/path-resolution.js';

export function resolveLauncherMainConfigPath(): string {
  return resolveConfigFilePath({
    appDataDir: process.env.APPDATA,
    xdgConfigHome: process.env.XDG_CONFIG_HOME,
    homeDir: os.homedir(),
    existsSync: fs.existsSync,
  });
}

export function readLauncherMainConfigObject(): Record<string, unknown> | null {
  const configPath = resolveLauncherMainConfigPath();
  if (!fs.existsSync(configPath)) return null;
  try {
    const data = fs.readFileSync(configPath, 'utf8');
    const parsed = configPath.endsWith('.jsonc') ? parseJsonc(data) : JSON.parse(data);
    if (!parsed || typeof parsed !== 'object') return null;
    return parsed as Record<string, unknown>;
  } catch {
    return null;
  }
}
