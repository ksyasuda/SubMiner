export function createCreateFirstRunSetupWindowHandler<TWindow>(deps: {
  createBrowserWindow: (options: Electron.BrowserWindowConstructorOptions) => TWindow;
}) {
  return (): TWindow =>
    deps.createBrowserWindow({
      width: 480,
      height: 460,
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
}

export function createCreateJellyfinSetupWindowHandler<TWindow>(deps: {
  createBrowserWindow: (options: Electron.BrowserWindowConstructorOptions) => TWindow;
}) {
  return (): TWindow =>
    deps.createBrowserWindow({
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
}

export function createCreateAnilistSetupWindowHandler<TWindow>(deps: {
  createBrowserWindow: (options: Electron.BrowserWindowConstructorOptions) => TWindow;
}) {
  return (): TWindow =>
    deps.createBrowserWindow({
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
}
