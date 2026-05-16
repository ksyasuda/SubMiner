import { constants, accessSync, realpathSync } from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { autoUpdater as electronAutoUpdater } from 'electron-updater';
import type { UpdateChannel } from '../../../types/config';
import { compareSemverLike } from './release-assets';

export interface AppUpdateCheckResult {
  available: boolean;
  version: string;
  canUpdate: boolean;
}

export interface ElectronUpdaterLoggerLike {
  info?: (message: string, ...args: unknown[]) => void;
  debug?: (message: string, ...args: unknown[]) => void;
  warn?: (message: string, ...args: unknown[]) => void;
  error?: (message: string, ...args: unknown[]) => void;
}

export interface ElectronAutoUpdaterLike {
  autoDownload: boolean;
  allowPrerelease: boolean;
  allowDowngrade: boolean;
  logger?: ElectronUpdaterLoggerLike | null;
  checkForUpdates: () => Promise<{
    updateInfo?: {
      version?: string;
    };
  } | null>;
  downloadUpdate: () => Promise<unknown>;
  quitAndInstall: (isSilent?: boolean, isForceRunAfter?: boolean) => void;
}

export function resolveMacAppBundlePath(execPath: string): string | null {
  const marker = '.app/Contents/MacOS/';
  const markerIndex = execPath.indexOf(marker);
  if (markerIndex < 0) return null;
  return execPath.slice(0, markerIndex + '.app'.length);
}

function readMacCodeSignature(appBundlePath: string): string | null {
  const result = spawnSync('/usr/bin/codesign', ['-dv', '--verbose=4', appBundlePath], {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  if (result.error || result.status !== 0) return null;
  return `${result.stdout ?? ''}\n${result.stderr ?? ''}`;
}

function canWriteLinuxAppImage(appImagePath: string): boolean {
  try {
    accessSync(appImagePath, constants.W_OK);
    accessSync(path.dirname(appImagePath), constants.W_OK);
    return true;
  } catch {
    return false;
  }
}

function realpathOrOriginal(filePath: string): string {
  try {
    return realpathSync(filePath);
  } catch {
    return filePath;
  }
}

export function isKnownLinuxPackageManagedAppImage(appImagePath: string): boolean {
  return realpathOrOriginal(appImagePath) === '/opt/SubMiner/SubMiner.AppImage';
}

export function isNativeUpdaterSupported(options: {
  platform: NodeJS.Platform;
  isPackaged: boolean;
  execPath: string;
  env?: NodeJS.ProcessEnv;
  canWriteAppImage?: (appImagePath: string) => boolean;
  readCodeSignature?: (appBundlePath: string) => string | null;
  log?: (message: string) => void;
}): boolean {
  if (!options.isPackaged) {
    options.log?.('Skipping native updater because this build is not packaged.');
    return false;
  }
  if (options.platform === 'linux') {
    const appImagePath = options.env?.APPIMAGE?.trim();
    if (!appImagePath) {
      options.log?.('Skipping native Linux updater because APPIMAGE is not set.');
      return false;
    }
    if (isKnownLinuxPackageManagedAppImage(appImagePath)) {
      options.log?.(
        'Skipping native Linux updater because this AppImage is managed by the system package manager.',
      );
      return false;
    }
    if (!(options.canWriteAppImage ?? canWriteLinuxAppImage)(appImagePath)) {
      options.log?.('Skipping native Linux updater because the running AppImage is not writable.');
      return false;
    }
    return true;
  }
  if (options.platform !== 'darwin') {
    options.log?.('Skipping native updater because this platform uses GitHub metadata checks.');
    return false;
  }

  const appBundlePath = resolveMacAppBundlePath(options.execPath);
  if (!appBundlePath) {
    options.log?.(
      'Skipping native macOS updater because the app bundle path could not be resolved.',
    );
    return false;
  }

  const signature = (options.readCodeSignature ?? readMacCodeSignature)(appBundlePath);
  if (!signature) {
    options.log?.(
      'Skipping native macOS updater because the app code signature could not be read.',
    );
    return false;
  }
  if (/Signature=adhoc\b/.test(signature) || /TeamIdentifier=not set\b/.test(signature)) {
    options.log?.('Skipping native macOS updater because this build is ad-hoc signed.');
    return false;
  }

  return true;
}

export function configureAutoUpdater(
  updater: ElectronAutoUpdaterLike,
  log: (message: string) => void = () => {},
  channel: UpdateChannel = 'stable',
): ElectronAutoUpdaterLike {
  updater.autoDownload = false;
  updater.allowPrerelease = channel === 'prerelease';
  updater.allowDowngrade = false;
  updater.logger = {
    info: () => {},
    debug: () => {},
    warn: (message) => log(message),
    error: (message) => log(message),
  };
  return updater;
}

export function createElectronAppUpdater(options: {
  currentVersion: string;
  isPackaged: boolean;
  updater?: ElectronAutoUpdaterLike;
  log: (message: string) => void;
  getChannel?: () => UpdateChannel;
  isNativeUpdaterSupported?: () => boolean;
}) {
  const getChannel = options.getChannel ?? (() => 'stable' as const);
  const updater = configureAutoUpdater(
    options.updater ?? electronAutoUpdater,
    options.log,
    getChannel(),
  );
  let nativeUpdaterSupported: boolean | null = null;

  function isNativeUpdaterSupported(): boolean {
    if (!options.isNativeUpdaterSupported) return true;
    if (nativeUpdaterSupported === null) {
      nativeUpdaterSupported = options.isNativeUpdaterSupported();
    }
    return nativeUpdaterSupported;
  }

  return {
    async checkForUpdates(channel?: UpdateChannel): Promise<AppUpdateCheckResult> {
      if (!options.isPackaged) {
        return {
          available: false,
          version: options.currentVersion,
          canUpdate: false,
        };
      }
      if (!isNativeUpdaterSupported()) {
        options.log('Skipping native app update check because native updater is unsupported.');
        return {
          available: false,
          version: options.currentVersion,
          canUpdate: false,
        };
      }
      configureAutoUpdater(updater, options.log, channel ?? getChannel());
      const result = await updater.checkForUpdates();
      const version = result?.updateInfo?.version ?? options.currentVersion;
      return {
        available: compareSemverLike(version, options.currentVersion) > 0,
        version,
        canUpdate: true,
      };
    },
    async downloadUpdate(): Promise<void> {
      if (!options.isPackaged) {
        options.log('Skipping app update download because this build is not packaged.');
        return;
      }
      if (!isNativeUpdaterSupported()) {
        options.log('Skipping app update download because native updater is unsupported.');
        return;
      }
      await updater.downloadUpdate();
    },
    quitAndInstall(): void {
      if (!options.isPackaged) {
        options.log('Skipping app update install because this build is not packaged.');
        return;
      }
      if (!isNativeUpdaterSupported()) {
        options.log('Skipping app update install because native updater is unsupported.');
        return;
      }
      updater.quitAndInstall(false, true);
    },
  };
}
