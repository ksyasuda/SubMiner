import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createBuildTrayMenuTemplateHandler,
  createResolveTrayIconPathHandler,
  shouldShowTexthookerTrayEntry,
} from './tray-main-actions';

test('resolve tray icon path handler forwards runtime dependencies', () => {
  const calls: string[] = [];
  const resolveTrayIconPath = createResolveTrayIconPathHandler({
    resolveTrayIconPathRuntime: (options) => {
      calls.push(`platform:${options.platform}`);
      calls.push(`resources:${options.resourcesPath}`);
      calls.push(`app:${options.appPath}`);
      calls.push(`dir:${options.dirname}`);
      calls.push(`join:${options.joinPath('a', 'b')}`);
      calls.push(`exists:${options.fileExists('/tmp/icon.png')}`);
      return '/tmp/icon.png';
    },
    platform: 'darwin',
    resourcesPath: '/resources',
    appPath: '/app',
    dirname: '/dir',
    joinPath: (...parts) => parts.join('/'),
    fileExists: () => true,
  });

  assert.equal(resolveTrayIconPath(), '/tmp/icon.png');
  assert.deepEqual(calls, [
    'platform:darwin',
    'resources:/resources',
    'app:/app',
    'dir:/dir',
    'join:a/b',
    'exists:true',
  ]);
});

test('build tray template handler wires actions and init guards', () => {
  const calls: string[] = [];
  let initialized = false;
  const buildTemplate = createBuildTrayMenuTemplateHandler({
    buildTrayMenuTemplateRuntime: (handlers) => {
      calls.push(`platform:${handlers.platform}`);
      handlers.openSessionHelp();
      handlers.openTexthookerInBrowser();
      calls.push(`show-texthooker:${handlers.showTexthookerPage}`);
      handlers.openFirstRunSetup();
      handlers.openWindowsMpvLauncherSetup();
      handlers.openYomitanSettings();
      handlers.openConfigSettings();
      handlers.openSyncUi();
      handlers.exportLogs();
      handlers.openJellyfinSetup();
      handlers.toggleJellyfinDiscovery(true);
      handlers.openAnilistSetup();
      handlers.checkForUpdates();
      handlers.quitApp();
      return [{ label: 'ok' }] as never;
    },
    initializeOverlayRuntime: () => {
      initialized = true;
      calls.push('init');
    },
    isOverlayRuntimeInitialized: () => initialized,
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
    toggleJellyfinDiscovery: async (checked) => {
      calls.push(`jellyfin-discovery:${checked}`);
    },
    platform: 'linux',
    openAnilistSetupWindow: () => calls.push('anilist'),
    checkForUpdates: () => calls.push('updates'),
    quitApp: () => calls.push('quit'),
  });

  const template = buildTemplate();
  assert.deepEqual(template, [{ label: 'ok' }]);
  assert.deepEqual(calls, [
    'platform:linux',
    'init',
    'help',
    'texthooker',
    'show-texthooker:true',
    'setup',
    'setup-forced',
    'yomitan',
    'configuration',
    'sync-ui',
    'export-logs',
    'jellyfin',
    'jellyfin-discovery:true',
    'anilist',
    'updates',
    'quit',
  ]);
});

test('windows mpv launcher tray action force-opens completed setup', () => {
  const calls: string[] = [];
  const buildTemplate = createBuildTrayMenuTemplateHandler({
    buildTrayMenuTemplateRuntime: (handlers) => {
      assert.equal(handlers.showFirstRunSetup, false);
      assert.equal(handlers.showWindowsMpvLauncherSetup, true);
      handlers.openWindowsMpvLauncherSetup();
      return [{ label: 'ok' }] as never;
    },
    initializeOverlayRuntime: () => calls.push('init'),
    isOverlayRuntimeInitialized: () => true,
    openSessionHelpModal: () => calls.push('help'),
    openChangelogModal: () => calls.push('changelog'),
    openTexthookerInBrowser: () => calls.push('texthooker'),
    showTexthookerPage: () => true,
    showFirstRunSetup: () => false,
    openFirstRunSetupWindow: (force?: boolean) => calls.push(force ? 'setup-forced' : 'setup'),
    showWindowsMpvLauncherSetup: () => true,
    openYomitanSettings: () => calls.push('yomitan'),
    openConfigSettingsWindow: () => calls.push('configuration'),
    openSyncUiWindow: () => calls.push('configuration'),
    exportLogs: () => calls.push('export-logs'),
    openJellyfinSetupWindow: () => calls.push('jellyfin'),
    isJellyfinConfigured: () => false,
    isJellyfinDiscoveryActive: () => false,
    toggleJellyfinDiscovery: () => {
      calls.push('jellyfin-discovery');
    },
    platform: 'win32',
    openAnilistSetupWindow: () => calls.push('anilist'),
    checkForUpdates: () => calls.push('updates'),
    quitApp: () => calls.push('quit'),
  });

  assert.deepEqual(buildTemplate(), [{ label: 'ok' }]);
  assert.deepEqual(calls, ['setup-forced']);
});

test('texthooker tray visibility follows websocket server enabled state', () => {
  assert.equal(
    shouldShowTexthookerTrayEntry({
      websocket: { enabled: false },
      annotationWebsocket: { enabled: false },
    }),
    false,
  );
  assert.equal(
    shouldShowTexthookerTrayEntry({
      websocket: { enabled: true },
      annotationWebsocket: { enabled: false },
    }),
    true,
  );
  assert.equal(
    shouldShowTexthookerTrayEntry({
      websocket: { enabled: 'auto' },
      annotationWebsocket: { enabled: false },
    }),
    true,
  );
  assert.equal(
    shouldShowTexthookerTrayEntry({
      websocket: { enabled: false },
      annotationWebsocket: { enabled: true },
    }),
    true,
  );
});
