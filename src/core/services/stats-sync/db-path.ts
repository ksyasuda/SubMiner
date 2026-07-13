import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { resolveConfigFilePath } from '../../../config/path-resolution';
import { parseConfigContent } from '../../../config/parse';
import { getDefaultConfigDir } from '../../../shared/setup-state';

/**
 * Default immersion stats database location, shared by the launcher's history
 * command and the app's --sync-cli mode: honor a configured
 * immersionTracking.dbPath (raw main-config read, tolerant of comments), else
 * <configDir>/immersion.sqlite. Electron- and libsql-free on purpose so the
 * bun launcher can import it.
 */
export function resolveImmersionDbPath(): string {
  const configPath = resolveConfigFilePath({
    appDataDir: process.env.APPDATA,
    xdgConfigHome: process.env.XDG_CONFIG_HOME,
    homeDir: os.homedir(),
    existsSync: fs.existsSync,
  });
  let configured = '';
  try {
    const parsed = parseConfigContent(configPath, fs.readFileSync(configPath, 'utf8'));
    const tracking =
      parsed && typeof parsed === 'object' && !Array.isArray(parsed)
        ? (parsed as Record<string, unknown>).immersionTracking
        : null;
    if (tracking && typeof tracking === 'object' && !Array.isArray(tracking)) {
      const dbPath = (tracking as Record<string, unknown>).dbPath;
      if (typeof dbPath === 'string') configured = dbPath.trim();
    }
  } catch {
    // no config or unreadable config → default location
  }
  if (configured) {
    return configured.startsWith('~')
      ? path.join(os.homedir(), configured.slice(1))
      : configured;
  }
  return path.join(getDefaultConfigDir(), 'immersion.sqlite');
}
