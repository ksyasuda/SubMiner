import fs from 'node:fs';
import os from 'node:os';
import { resolveConfigFilePath } from '../src/config/path-resolution.js';

export function resolveMainConfigPath(): string {
  return resolveConfigFilePath({
    xdgConfigHome: process.env.XDG_CONFIG_HOME,
    homeDir: os.homedir(),
    existsSync: fs.existsSync,
  });
}
