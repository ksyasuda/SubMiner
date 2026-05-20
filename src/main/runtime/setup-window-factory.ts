interface SetupWindowConfig {
  width: number;
  height: number;
  title: string;
  resizable?: boolean;
  minimizable?: boolean;
  maximizable?: boolean;
  preloadPath?: string;
  sandbox?: boolean;
  backgroundColor?: string;
}

function createSetupWindowHandler<TWindow>(
  deps: { createBrowserWindow: (options: Electron.BrowserWindowConstructorOptions) => TWindow },
  config: SetupWindowConfig,
) {
  return (): TWindow =>
    deps.createBrowserWindow({
      width: config.width,
      height: config.height,
      title: config.title,
      show: true,
      autoHideMenuBar: true,
      ...(config.resizable === undefined ? {} : { resizable: config.resizable }),
      ...(config.minimizable === undefined ? {} : { minimizable: config.minimizable }),
      ...(config.maximizable === undefined ? {} : { maximizable: config.maximizable }),
      ...(config.backgroundColor === undefined ? {} : { backgroundColor: config.backgroundColor }),
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
        ...(config.sandbox === undefined ? {} : { sandbox: config.sandbox }),
        ...(config.preloadPath ? { preload: config.preloadPath } : {}),
      },
    });
}

export function createCreateFirstRunSetupWindowHandler<TWindow>(deps: {
  createBrowserWindow: (options: Electron.BrowserWindowConstructorOptions) => TWindow;
}) {
  return createSetupWindowHandler(deps, {
    width: 720,
    height: 860,
    title: 'SubMiner Setup',
    resizable: false,
    minimizable: false,
    maximizable: false,
  });
}

export function createCreateJellyfinSetupWindowHandler<TWindow>(deps: {
  createBrowserWindow: (options: Electron.BrowserWindowConstructorOptions) => TWindow;
  preloadPath?: string;
}) {
  return createSetupWindowHandler(deps, {
    width: 520,
    height: 560,
    title: 'Jellyfin Setup',
    ...(deps.preloadPath ? { preloadPath: deps.preloadPath, sandbox: true } : {}),
  });
}

export function createCreateAnilistSetupWindowHandler<TWindow>(deps: {
  createBrowserWindow: (options: Electron.BrowserWindowConstructorOptions) => TWindow;
}) {
  return createSetupWindowHandler(deps, {
    width: 1000,
    height: 760,
    title: 'Anilist Setup',
  });
}

export function createCreateConfigSettingsWindowHandler<TWindow>(deps: {
  createBrowserWindow: (options: Electron.BrowserWindowConstructorOptions) => TWindow;
  preloadPath: string;
}) {
  return createSetupWindowHandler(deps, {
    width: 1040,
    height: 760,
    title: 'SubMiner Configuration',
    resizable: true,
    preloadPath: deps.preloadPath,
    sandbox: false,
    backgroundColor: '#24273a',
  });
}
