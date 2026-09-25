import fs from 'node:fs';
import path from 'node:path';
import { cleanupOldWindowsManagedRuntimes, stageManagedLauncher } from './managed-launcher';

// Bundled separately as Node-compatible code so AppImages can prepare their
// runtime without starting Electron's GUI, single-instance lock, or settings.
export function prepareLauncherRuntime(options: { appPath: string; resourcesPath: string }) {
  const appVersion = fs
    .readFileSync(path.join(options.resourcesPath, 'launcher', 'version'), 'utf8')
    .trim();
  const payload = stageManagedLauncher({
    appExePath: options.appPath,
    appVersion,
    bundledBunPath: path.join(
      options.resourcesPath,
      'bun',
      process.platform === 'win32' ? 'bun.exe' : 'bun',
    ),
    launcherResourcePath: path.join(options.resourcesPath, 'launcher', 'subminer.js'),
    force: process.platform === 'linux',
  });
  if (process.platform === 'win32') cleanupOldWindowsManagedRuntimes({ appVersion });
  return payload;
}
