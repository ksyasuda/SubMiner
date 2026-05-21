import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createCreateAnilistSetupWindowHandler,
  createCreateConfigSettingsWindowHandler,
  createCreateFirstRunSetupWindowHandler,
  createCreateJellyfinSetupWindowHandler,
} from './setup-window-factory';

test('createCreateFirstRunSetupWindowHandler builds first-run setup window', () => {
  let options: Electron.BrowserWindowConstructorOptions | null = null;
  const createSetupWindow = createCreateFirstRunSetupWindowHandler({
    createBrowserWindow: (nextOptions) => {
      options = nextOptions;
      return { id: 'first-run' } as never;
    },
  });

  assert.deepEqual(createSetupWindow(), { id: 'first-run' });
  assert.deepEqual(options, {
    width: 720,
    height: 860,
    title: 'SubMiner Setup',
    show: true,
    autoHideMenuBar: true,
    resizable: false,
    minimizable: false,
    maximizable: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
    },
  });
});

test('createCreateJellyfinSetupWindowHandler builds jellyfin setup window', () => {
  let options: Electron.BrowserWindowConstructorOptions | null = null;
  const createSetupWindow = createCreateJellyfinSetupWindowHandler({
    createBrowserWindow: (nextOptions) => {
      options = nextOptions;
      return { id: 'jellyfin' } as never;
    },
  });

  assert.deepEqual(createSetupWindow(), { id: 'jellyfin' });
  assert.deepEqual(options, {
    width: 520,
    height: 560,
    title: 'Jellyfin Setup',
    show: true,
    autoHideMenuBar: true,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
    },
  });
});

test('createCreateJellyfinSetupWindowHandler wires optional preload bridge', () => {
  const captured: { options?: Electron.BrowserWindowConstructorOptions } = {};
  const createSetupWindow = createCreateJellyfinSetupWindowHandler({
    createBrowserWindow: (nextOptions) => {
      captured.options = nextOptions;
      return { id: 'jellyfin' } as never;
    },
    preloadPath: 'C:\\SubMiner\\dist\\preload-jellyfin-setup.js',
  });

  assert.deepEqual(createSetupWindow(), { id: 'jellyfin' });
  const options = captured.options;
  assert.ok(options);
  assert.equal(options.webPreferences?.preload, 'C:\\SubMiner\\dist\\preload-jellyfin-setup.js');
  assert.equal(options.webPreferences?.sandbox, true);
});

test('createCreateAnilistSetupWindowHandler builds anilist setup window', () => {
  let options: Electron.BrowserWindowConstructorOptions | null = null;
  const createSetupWindow = createCreateAnilistSetupWindowHandler({
    createBrowserWindow: (nextOptions) => {
      options = nextOptions;
      return { id: 'anilist' } as never;
    },
  });

  assert.deepEqual(createSetupWindow(), { id: 'anilist' });
  assert.deepEqual(options, {
    width: 1000,
    height: 760,
    title: 'Anilist Setup',
    show: true,
    autoHideMenuBar: true,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
    },
  });
});

test('createCreateConfigSettingsWindowHandler builds configuration settings window', () => {
  let options: Electron.BrowserWindowConstructorOptions | null = null;
  const createSettingsWindow = createCreateConfigSettingsWindowHandler({
    preloadPath: '/tmp/preload-settings.js',
    createBrowserWindow: (nextOptions) => {
      options = nextOptions;
      return { id: 'config-settings' } as never;
    },
  });

  assert.deepEqual(createSettingsWindow(), { id: 'config-settings' });
  assert.deepEqual(options, {
    width: 1040,
    height: 760,
    title: 'SubMiner Settings',
    show: true,
    autoHideMenuBar: true,
    resizable: true,
    backgroundColor: '#24273a',
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: '/tmp/preload-settings.js',
    },
  });
});
