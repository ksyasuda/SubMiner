interface SetupWindowConfig {
  width: number;
  height: number;
  title: string;
  resizable?: boolean;
  minimizable?: boolean;
  maximizable?: boolean;
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
      resizable: config.resizable,
      minimizable: config.minimizable,
      maximizable: config.maximizable,
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
      },
    });
}

export function createCreateFirstRunSetupWindowHandler<TWindow>(deps: {
  createBrowserWindow: (options: Electron.BrowserWindowConstructorOptions) => TWindow;
}) {
  return createSetupWindowHandler(deps, {
    width: 480,
    height: 460,
    title: 'SubMiner Setup',
    resizable: false,
    minimizable: false,
    maximizable: false,
  });
}

export function createCreateJellyfinSetupWindowHandler<TWindow>(deps: {
  createBrowserWindow: (options: Electron.BrowserWindowConstructorOptions) => TWindow;
}) {
  return createSetupWindowHandler(deps, {
    width: 520,
    height: 560,
    title: 'Jellyfin Setup',
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
