import test from 'node:test';
import assert from 'node:assert/strict';
import {
  configureAutoUpdater,
  createElectronAppUpdater,
  isKnownLinuxPackageManagedAppImage,
  isNativeUpdaterSupported,
  resolveMacAppBundlePath,
  type ElectronAutoUpdaterLike,
} from './app-updater';

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

test('app updater skips native update checks when native updater is unsupported', async () => {
  let checked = false;
  const updater: ElectronAutoUpdaterLike = {
    autoDownload: true,
    allowPrerelease: false,
    allowDowngrade: true,
    logger: null,
    checkForUpdates: async () => {
      checked = true;
      return {
        updateInfo: {
          version: '0.15.0',
        },
      };
    },
    downloadUpdate: async () => [],
    quitAndInstall: () => {},
  };
  const logged: string[] = [];
  const appUpdater = createElectronAppUpdater({
    currentVersion: '0.14.0',
    isPackaged: true,
    updater,
    log: (message) => logged.push(message),
    isNativeUpdaterSupported: () => false,
  });

  const result = await appUpdater.checkForUpdates('stable');

  assert.equal(checked, false);
  assert.deepEqual(result, {
    available: false,
    version: '0.14.0',
    canUpdate: false,
  });
  assert.deepEqual(logged, [
    'Skipping native app update check because native updater is unsupported.',
  ]);
});

test('app updater skips native downloads when native updater is unsupported', async () => {
  let downloaded = false;
  const updater: ElectronAutoUpdaterLike = {
    autoDownload: true,
    allowPrerelease: false,
    allowDowngrade: true,
    logger: null,
    checkForUpdates: async () => null,
    downloadUpdate: async () => {
      downloaded = true;
      return [];
    },
    quitAndInstall: () => {},
  };
  const logged: string[] = [];
  const appUpdater = createElectronAppUpdater({
    currentVersion: '0.14.0',
    isPackaged: true,
    updater,
    log: (message) => logged.push(message),
    isNativeUpdaterSupported: () => false,
  });

  await appUpdater.downloadUpdate();

  assert.equal(downloaded, false);
  assert.deepEqual(logged, ['Skipping app update download because native updater is unsupported.']);
});

test('resolveMacAppBundlePath resolves packaged macOS executable path', () => {
  assert.equal(
    resolveMacAppBundlePath('/Applications/SubMiner.app/Contents/MacOS/SubMiner'),
    '/Applications/SubMiner.app',
  );
  assert.equal(resolveMacAppBundlePath('/usr/local/bin/SubMiner'), null);
});

test('mac native updater is unsupported for ad-hoc signed app bundles', () => {
  const logged: string[] = [];
  const supported = isNativeUpdaterSupported({
    platform: 'darwin',
    isPackaged: true,
    execPath: '/Applications/SubMiner.app/Contents/MacOS/SubMiner',
    readCodeSignature: () =>
      ['Signature=adhoc', 'TeamIdentifier=not set', 'Runtime Version=26.0.0'].join('\n'),
    log: (message) => logged.push(message),
  });

  assert.equal(supported, false);
  assert.deepEqual(logged, ['Skipping native macOS updater because this build is ad-hoc signed.']);
});

test('mac native updater is supported for Developer ID signed app bundles', () => {
  const supported = isNativeUpdaterSupported({
    platform: 'darwin',
    isPackaged: true,
    execPath: '/Applications/SubMiner.app/Contents/MacOS/SubMiner',
    readCodeSignature: () =>
      ['Authority=Developer ID Application: Example', 'TeamIdentifier=ABCDE12345'].join('\n'),
  });

  assert.equal(supported, true);
});

test('linux native updater is supported for writable direct AppImage installs', () => {
  const supported = isNativeUpdaterSupported({
    platform: 'linux',
    isPackaged: true,
    execPath: '/tmp/.mount_SubMiner/SubMiner',
    env: {
      APPIMAGE: '/home/tester/.local/bin/SubMiner.AppImage',
    },
    canWriteAppImage: (appImagePath) =>
      appImagePath === '/home/tester/.local/bin/SubMiner.AppImage',
  });

  assert.equal(supported, true);
});

test('linux native updater is unsupported when APPIMAGE is missing', () => {
  const logged: string[] = [];
  const supported = isNativeUpdaterSupported({
    platform: 'linux',
    isPackaged: true,
    execPath: '/tmp/.mount_SubMiner/SubMiner',
    env: {},
    log: (message) => logged.push(message),
  });

  assert.equal(supported, false);
  assert.deepEqual(logged, ['Skipping native Linux updater because APPIMAGE is not set.']);
});

test('linux native updater is unsupported for non-writable AppImage installs', () => {
  const logged: string[] = [];
  const supported = isNativeUpdaterSupported({
    platform: 'linux',
    isPackaged: true,
    execPath: '/tmp/.mount_SubMiner/SubMiner',
    env: {
      APPIMAGE: '/home/tester/.local/bin/SubMiner.AppImage',
    },
    canWriteAppImage: () => false,
    log: (message) => logged.push(message),
  });

  assert.equal(supported, false);
  assert.deepEqual(logged, [
    'Skipping native Linux updater because the running AppImage is not writable.',
  ]);
});

test('linux native updater is unsupported for package-managed AppImage installs', () => {
  const logged: string[] = [];
  const supported = isNativeUpdaterSupported({
    platform: 'linux',
    isPackaged: true,
    execPath: '/tmp/.mount_SubMiner/SubMiner',
    env: {
      APPIMAGE: '/opt/SubMiner/SubMiner.AppImage',
    },
    canWriteAppImage: () => true,
    log: (message) => logged.push(message),
  });

  assert.equal(supported, false);
  assert.deepEqual(logged, [
    'Skipping native Linux updater because this AppImage is managed by the system package manager.',
  ]);
});

test('known Linux package-managed AppImage detection follows the canonical AUR path', () => {
  assert.equal(isKnownLinuxPackageManagedAppImage('/opt/SubMiner/SubMiner.AppImage'), true);
  assert.equal(
    isKnownLinuxPackageManagedAppImage('/home/tester/.local/bin/SubMiner.AppImage'),
    false,
  );
});

test('native updater is unsupported on Windows by default', () => {
  const supported = isNativeUpdaterSupported({
    platform: 'win32',
    isPackaged: true,
    execPath: 'C:\\Program Files\\SubMiner\\SubMiner.exe',
  });

  assert.equal(supported, false);
});
