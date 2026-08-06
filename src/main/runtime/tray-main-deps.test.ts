import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildResolveTrayIconPathMainDepsHandler,
  createBuildTrayMenuTemplateMainDepsHandler,
} from './tray-main-deps';

test('tray main deps builders return mapped handlers', () => {
  const calls: string[] = [];
  const resolveDeps = createBuildResolveTrayIconPathMainDepsHandler({
    resolveTrayIconPathRuntime: () => '/tmp/icon.png',
    platform: 'darwin',
    resourcesPath: '/resources',
    appPath: '/app',
    dirname: '/dir',
    joinPath: (...parts) => parts.join('/'),
    fileExists: () => true,
  })();

  assert.equal(resolveDeps.platform, 'darwin');
  assert.equal(resolveDeps.joinPath('a', 'b'), 'a/b');

  const menuDeps = createBuildTrayMenuTemplateMainDepsHandler({
    buildTrayMenuTemplateRuntime: () => [{ label: 'tray' }] as never,
    initializeOverlayRuntime: () => calls.push('init'),
    isOverlayRuntimeInitialized: () => false,
    openSessionHelpModal: () => calls.push('help'),
    openChangelogModal: () => calls.push('changelog'),
    openTexthookerInBrowser: () => calls.push('texthooker'),
    showTexthookerPage: () => true,
    showFirstRunSetup: () => true,
    openFirstRunSetupWindow: (force?: boolean) => calls.push(force ? 'setup-forced' : 'setup'),
    showWindowsMpvLauncherSetup: () => true,
    openYomitanSettings: () => calls.push('yomitan'),
    openConfigSettingsWindow: () => calls.push('configuration'),
    openSyncUiWindow: () => calls.push('sync-ui'),
    exportLogs: () => calls.push('export-logs'),
    openJellyfinSetupWindow: () => calls.push('jellyfin'),
    isJellyfinConfigured: () => true,
    isJellyfinDiscoveryActive: () => false,
    toggleJellyfinDiscovery: (checked) => {
      calls.push(`jellyfin-discovery:${checked}`);
    },
    platform: 'linux',
    openAnilistSetupWindow: () => calls.push('anilist'),
    checkForUpdates: () => calls.push('updates'),
    quitApp: () => calls.push('quit'),
  })();

  assert.equal(menuDeps.platform, 'linux');
  const template = menuDeps.buildTrayMenuTemplateRuntime({
    platform: menuDeps.platform,
    openSessionHelp: () => calls.push('open-help'),
    openChangelog: () => calls.push('open-changelog'),
    openTexthookerInBrowser: () => calls.push('open-texthooker'),
    showTexthookerPage: true,
    openFirstRunSetup: () => calls.push('open-setup'),
    showFirstRunSetup: true,
    openWindowsMpvLauncherSetup: () => calls.push('open-windows-mpv'),
    showWindowsMpvLauncherSetup: true,
    openYomitanSettings: () => calls.push('open-yomitan'),
    openConfigSettings: () => calls.push('open-configuration'),
    openSyncUi: () => {},
    exportLogs: () => calls.push('open-export-logs'),
    openJellyfinSetup: () => calls.push('open-jellyfin'),
    showJellyfinDiscovery: true,
    jellyfinDiscoveryActive: false,
    toggleJellyfinDiscovery: (checked) => calls.push(`open-jellyfin-discovery:${checked}`),
    openAnilistSetup: () => calls.push('open-anilist'),
    checkForUpdates: () => calls.push('open-updates'),
    quitApp: () => calls.push('quit-app'),
  });

  assert.deepEqual(template, [{ label: 'tray' }]);
});
