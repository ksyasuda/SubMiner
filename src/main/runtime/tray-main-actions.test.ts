import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createBuildTrayMenuTemplateHandler,
  createResolveTrayIconPathHandler,
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
      handlers.openSessionHelp();
      handlers.openTexthookerInBrowser();
      handlers.openFirstRunSetup();
      handlers.openWindowsMpvLauncherSetup();
      handlers.openYomitanSettings();
      handlers.openRuntimeOptions();
      handlers.openJellyfinSetup();
      handlers.openAnilistSetup();
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
    showFirstRunSetup: () => true,
    openFirstRunSetupWindow: () => calls.push('setup'),
    showWindowsMpvLauncherSetup: () => true,
    openYomitanSettings: () => calls.push('yomitan'),
    openRuntimeOptionsPalette: () => calls.push('runtime-options'),
    openJellyfinSetupWindow: () => calls.push('jellyfin'),
    openAnilistSetupWindow: () => calls.push('anilist'),
    quitApp: () => calls.push('quit'),
  });

  const template = buildTemplate();
  assert.deepEqual(template, [{ label: 'ok' }]);
  assert.deepEqual(calls, [
    'init',
    'help',
    'texthooker',
    'setup',
    'setup',
    'yomitan',
    'runtime-options',
    'jellyfin',
    'anilist',
    'quit',
  ]);
});
