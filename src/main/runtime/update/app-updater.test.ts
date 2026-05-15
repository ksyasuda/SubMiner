import test from 'node:test';
import assert from 'node:assert/strict';
import { configureAutoUpdater, type ElectronAutoUpdaterLike } from './app-updater';

type UpdaterLogger = {
  info: (message: string) => void;
  debug: (message: string) => void;
  warn: (message: string) => void;
  error: (message: string) => void;
};

test('configureAutoUpdater disables eager update behavior and suppresses info logging', () => {
  const logged: string[] = [];
  const updater: ElectronAutoUpdaterLike & { logger?: UpdaterLogger | null } = {
    autoDownload: true,
    allowPrerelease: true,
    allowDowngrade: true,
    logger: null,
    checkForUpdates: async () => null,
    downloadUpdate: async () => [],
    quitAndInstall: () => {},
  };

  configureAutoUpdater(updater, (message) => logged.push(message));

  assert.equal(updater.autoDownload, false);
  assert.equal(updater.allowPrerelease, false);
  assert.equal(updater.allowDowngrade, false);
  assert.ok(updater.logger);

  updater.logger.info('Checking for update');
  updater.logger.debug('Generated new staging user ID');
  updater.logger.warn('metadata missing');
  updater.logger.error('download failed');

  assert.deepEqual(logged, ['metadata missing', 'download failed']);
});

test('configureAutoUpdater allows prereleases only for the prerelease channel', () => {
  const updater: ElectronAutoUpdaterLike = {
    autoDownload: true,
    allowPrerelease: false,
    allowDowngrade: true,
    logger: null,
    checkForUpdates: async () => null,
    downloadUpdate: async () => [],
    quitAndInstall: () => {},
  };

  configureAutoUpdater(updater, () => {}, 'prerelease');
  assert.equal(updater.allowPrerelease, true);

  configureAutoUpdater(updater, () => {}, 'stable');
  assert.equal(updater.allowPrerelease, false);
});
