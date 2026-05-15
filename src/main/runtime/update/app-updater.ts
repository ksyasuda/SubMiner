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
}) {
  const getChannel = options.getChannel ?? (() => 'stable' as const);
  const updater = configureAutoUpdater(
    options.updater ?? electronAutoUpdater,
    options.log,
    getChannel(),
  );

  return {
    async checkForUpdates(channel?: UpdateChannel): Promise<AppUpdateCheckResult> {
      if (!options.isPackaged) {
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
      await updater.downloadUpdate();
    },
    quitAndInstall(): void {
      updater.quitAndInstall(false, true);
    },
  };
}
