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
    openTexthookerInBrowser: () => calls.push('texthooker'),
    showTexthookerPage: () => true,
    showFirstRunSetup: () => true,
    openFirstRunSetupWindow: () => calls.push('setup'),
    showWindowsMpvLauncherSetup: () => true,
    openYomitanSettings: () => calls.push('yomitan'),
    openRuntimeOptionsPalette: () => calls.push('runtime-options'),
    openJellyfinSetupWindow: () => calls.push('jellyfin'),
    isJellyfinConfigured: () => true,
    isJellyfinDiscoveryActive: () => false,
    toggleJellyfinDiscovery: () => {
      calls.push('jellyfin-discovery');
    },
    openAnilistSetupWindow: () => calls.push('anilist'),
    checkForUpdates: () => calls.push('updates'),
    quitApp: () => calls.push('quit'),
  })();

  const template = menuDeps.buildTrayMenuTemplateRuntime({
    openSessionHelp: () => calls.push('open-help'),
    openTexthookerInBrowser: () => calls.push('open-texthooker'),
    showTexthookerPage: true,
    openFirstRunSetup: () => calls.push('open-setup'),
    showFirstRunSetup: true,
    openWindowsMpvLauncherSetup: () => calls.push('open-windows-mpv'),
    showWindowsMpvLauncherSetup: true,
    openYomitanSettings: () => calls.push('open-yomitan'),
    openRuntimeOptions: () => calls.push('open-runtime-options'),
    openJellyfinSetup: () => calls.push('open-jellyfin'),
    showJellyfinDiscovery: true,
    jellyfinDiscoveryActive: false,
    toggleJellyfinDiscovery: () => calls.push('open-jellyfin-discovery'),
    openAnilistSetup: () => calls.push('open-anilist'),
    checkForUpdates: () => calls.push('open-updates'),
    quitApp: () => calls.push('quit-app'),
  });

  assert.deepEqual(template, [{ label: 'tray' }]);
});
