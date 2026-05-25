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
    openTexthookerInBrowser: () => calls.push('texthooker'),
    showTexthookerPage: () => true,
    showFirstRunSetup: () => true,
    openFirstRunSetupWindow: () => calls.push('setup'),
    showWindowsMpvLauncherSetup: () => true,
    openYomitanSettings: () => calls.push('yomitan'),
    openConfigSettingsWindow: () => calls.push('configuration'),
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
    'setup',
    'yomitan',
    'configuration',
    'jellyfin',
    'jellyfin-discovery:true',
    'anilist',
    'updates',
    'quit',
  ]);
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
