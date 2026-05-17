import { realpathSync } from 'node:fs';
import { execFile } from 'node:child_process';
import os from 'node:os';
import path from 'node:path';
import { promisify } from 'node:util';
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
  autoInstallOnAppQuit?: boolean;
  allowPrerelease: boolean;
  allowDowngrade: boolean;
  logger?: ElectronUpdaterLoggerLike | null;
  on?: (event: 'error', listener: (error: unknown) => void) => unknown;
  off?: (event: 'error', listener: (error: unknown) => void) => unknown;
  removeListener?: (event: 'error', listener: (error: unknown) => void) => unknown;
  checkForUpdates: () => Promise<{
    updateInfo?: {
      version?: string;
    };
  } | null>;
  downloadUpdate: () => Promise<unknown>;
  quitAndInstall: (isSilent?: boolean, isForceRunAfter?: boolean) => void;
  disableDifferentialDownload?: boolean;
}

const updaterErrorListeners = new WeakMap<object, (error: unknown) => void>();
const execFileAsync = promisify(execFile);

type ElectronAutoUpdaterWithHttpExecutor = ElectronAutoUpdaterLike & {
  httpExecutor?: unknown;
};

export function resolveMacAppBundlePath(execPath: string): string | null {
  const marker = '.app/Contents/MacOS/';
  const markerIndex = execPath.indexOf(marker);
  if (markerIndex < 0) return null;
  return execPath.slice(0, markerIndex + '.app'.length);
}

async function readMacCodeSignature(appBundlePath: string): Promise<string | null> {
  try {
    const result = await execFileAsync('/usr/bin/codesign', ['-dv', '--verbose=4', appBundlePath], {
      encoding: 'utf8',
    });
    return `${result.stdout ?? ''}\n${result.stderr ?? ''}`;
  } catch {
    return null;
  }
}

function realpathOrOriginal(filePath: string): string {
  try {
    return realpathSync(filePath);
  } catch {
    return filePath;
  }
}

function isSameOrInsideDirectory(parentPath: string, candidatePath: string): boolean {
  const relative = path.relative(parentPath, candidatePath);
  return (
    relative === '' ||
    (relative.length > 0 && !relative.startsWith('..') && !path.isAbsolute(relative))
  );
}

export function isMacApplicationsFolderBundle(
  appBundlePath: string,
  homeDir: string = os.homedir(),
): boolean {
  const resolvedBundlePath = path.resolve(appBundlePath);
  return (
    isSameOrInsideDirectory('/Applications', resolvedBundlePath) ||
    isSameOrInsideDirectory(path.join(homeDir, 'Applications'), resolvedBundlePath)
  );
}

export function isKnownLinuxPackageManagedAppImage(appImagePath: string): boolean {
  return realpathOrOriginal(appImagePath) === '/opt/SubMiner/SubMiner.AppImage';
}

export async function isNativeUpdaterSupported(options: {
  platform: NodeJS.Platform;
  isPackaged: boolean;
  execPath: string;
  env?: NodeJS.ProcessEnv;
  homeDir?: string;
  readCodeSignature?: (appBundlePath: string) => string | null | Promise<string | null>;
  log?: (message: string) => void;
}): Promise<boolean> {
  if (!options.isPackaged) {
    options.log?.('Skipping native updater because this build is not packaged.');
    return false;
  }
  if (options.platform === 'linux') {
    options.log?.(
      'Skipping native Linux updater because Linux tray checks use GitHub release assets.',
    );
    return false;
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

  if (!isMacApplicationsFolderBundle(appBundlePath, options.homeDir)) {
    options.log?.(
      'Skipping native macOS updater because the app is not installed in an Applications folder.',
    );
    return false;
  }

  const signature = await (options.readCodeSignature ?? readMacCodeSignature)(appBundlePath);
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
  // On macOS this avoids invoking Squirrel until the explicit restart/install step.
  updater.autoInstallOnAppQuit = false;
  updater.allowPrerelease = channel === 'prerelease';
  updater.allowDowngrade = false;
  updater.logger = {
    info: () => {},
    debug: () => {},
    warn: (message) => log(message),
    error: (message) => log(message),
  };
  const previousErrorListener = updaterErrorListeners.get(updater);
  if (previousErrorListener) {
    if (updater.off) {
      updater.off('error', previousErrorListener);
    } else {
      updater.removeListener?.('error', previousErrorListener);
    }
  }
  if (updater.on) {
    const errorListener = (error: unknown) => {
      const message = error instanceof Error ? error.message : String(error);
      log(`Updater error event: ${message}`);
    };
    updater.on('error', errorListener);
    updaterErrorListeners.set(updater, errorListener);
  }
  return updater;
}

export function createElectronAppUpdater(options: {
  currentVersion: string;
  isPackaged: boolean;
  updater?: ElectronAutoUpdaterLike;
  log: (message: string) => void;
  getChannel?: () => UpdateChannel;
  isNativeUpdaterSupported?: () => boolean | Promise<boolean>;
  configureHttpExecutor?: () => unknown;
  disableDifferentialDownload?: boolean;
}) {
  const getChannel = options.getChannel ?? (() => 'stable' as const);
  const updater = configureAutoUpdater(
    options.updater ?? electronAutoUpdater,
    options.log,
    getChannel(),
  );
  if (options.configureHttpExecutor) {
    // electron-updater has no public executor hook; keep the macOS cURL override localized.
    (updater as ElectronAutoUpdaterWithHttpExecutor).httpExecutor = options.configureHttpExecutor();
  }
  if (options.disableDifferentialDownload !== undefined) {
    updater.disableDifferentialDownload = options.disableDifferentialDownload;
  }
  let nativeUpdaterSupported: Promise<boolean> | null = null;

  async function getNativeUpdaterSupported(): Promise<boolean> {
    if (!options.isNativeUpdaterSupported) return true;
    if (nativeUpdaterSupported === null) {
      nativeUpdaterSupported = Promise.resolve(options.isNativeUpdaterSupported());
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
      if (!(await getNativeUpdaterSupported())) {
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
      if (!(await getNativeUpdaterSupported())) {
        options.log('Skipping app update download because native updater is unsupported.');
        return;
      }
      await updater.downloadUpdate();
    },
    async quitAndInstall(): Promise<void> {
      if (!options.isPackaged) {
        options.log('Skipping app update install because this build is not packaged.');
        return;
      }
      if (!(await getNativeUpdaterSupported())) {
        options.log('Skipping app update install because native updater is unsupported.');
        return;
      }
      updater.quitAndInstall(false, true);
    },
  };
}
