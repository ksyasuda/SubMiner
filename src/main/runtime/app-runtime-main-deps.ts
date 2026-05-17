export function createBuildEnsureTrayMainDepsHandler<TTray, TTrayMenu, TTrayIcon>(deps: {
  getTray: () => TTray | null;
  setTray: (tray: TTray | null) => void;
  buildTrayMenu: () => TTrayMenu;
  resolveTrayIconPath: () => string | null;
  createImageFromPath: (iconPath: string) => TTrayIcon;
  createEmptyImage: () => TTrayIcon;
  createTray: (icon: TTrayIcon) => TTray;
  trayTooltip: string;
  platform: string;
  logWarn: (message: string) => void;
  initializeOverlayRuntime: () => void;
  isOverlayRuntimeInitialized: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
}) {
  return () => ({
    getTray: () => deps.getTray(),
    setTray: (tray: TTray | null) => deps.setTray(tray),
    buildTrayMenu: () => deps.buildTrayMenu(),
    resolveTrayIconPath: () => deps.resolveTrayIconPath(),
    createImageFromPath: (iconPath: string) => deps.createImageFromPath(iconPath),
    createEmptyImage: () => deps.createEmptyImage(),
    createTray: (icon: TTrayIcon) => deps.createTray(icon),
    trayTooltip: deps.trayTooltip,
    platform: deps.platform,
    logWarn: (message: string) => deps.logWarn(message),
    ensureOverlayVisibleFromTrayClick: () => {
      if (!deps.isOverlayRuntimeInitialized()) {
        deps.initializeOverlayRuntime();
      }
      deps.setVisibleOverlayVisible(true);
    },
  });
}

export function createBuildDestroyTrayMainDepsHandler<TTray>(deps: {
  getTray: () => TTray | null;
  setTray: (tray: TTray | null) => void;
}) {
  return () => ({
    getTray: () => deps.getTray(),
    setTray: (tray: TTray | null) => deps.setTray(tray),
  });
}

export function createBuildInitializeOverlayRuntimeBootstrapMainDepsHandler<TOptions>(deps: {
  isOverlayRuntimeInitialized: () => boolean;
  initializeOverlayRuntimeCore: (options: TOptions) => void;
  buildOptions: () => TOptions;
  setOverlayRuntimeInitialized: (initialized: boolean) => void;
  startBackgroundWarmups: () => void;
}) {
  return () => ({
    isOverlayRuntimeInitialized: () => deps.isOverlayRuntimeInitialized(),
    initializeOverlayRuntimeCore: (options: TOptions) => deps.initializeOverlayRuntimeCore(options),
    buildOptions: () => deps.buildOptions(),
    setOverlayRuntimeInitialized: (initialized: boolean) =>
      deps.setOverlayRuntimeInitialized(initialized),
    startBackgroundWarmups: () => deps.startBackgroundWarmups(),
  });
}

export function createBuildOpenYomitanSettingsMainDepsHandler<TYomitanExt, TWindow>(deps: {
  ensureYomitanExtensionLoaded: () => Promise<TYomitanExt | null>;
  getYomitanExtension?: () => TYomitanExt | null;
  getYomitanExtensionLoadInFlight?: () => Promise<unknown> | null;
  openYomitanSettingsWindow: (params: {
    yomitanExt: TYomitanExt;
    getExistingWindow: () => TWindow | null;
    setWindow: (window: TWindow | null) => void;
    yomitanSession?: unknown | null;
    onWindowClosed?: () => void;
  }) => void;
  getExistingWindow: () => TWindow | null;
  setWindow: (window: TWindow | null) => void;
  getYomitanSession?: () => unknown | null;
  logWarn: (message: string) => void;
  logError: (message: string, error: unknown) => void;
}) {
  return () => ({
    ensureYomitanExtensionLoaded: () => deps.ensureYomitanExtensionLoaded(),
    ...(deps.getYomitanExtension
      ? { getYomitanExtension: () => deps.getYomitanExtension?.() ?? null }
      : {}),
    ...(deps.getYomitanExtensionLoadInFlight
      ? {
          getYomitanExtensionLoadInFlight: () => deps.getYomitanExtensionLoadInFlight?.() ?? null,
        }
      : {}),
    openYomitanSettingsWindow: (params: {
      yomitanExt: TYomitanExt;
      getExistingWindow: () => TWindow | null;
      setWindow: (window: TWindow | null) => void;
      yomitanSession?: unknown | null;
      onWindowClosed?: () => void;
    }) => deps.openYomitanSettingsWindow(params),
    getExistingWindow: () => deps.getExistingWindow(),
    setWindow: (window: TWindow | null) => deps.setWindow(window),
    getYomitanSession: () => deps.getYomitanSession?.() ?? null,
    logWarn: (message: string) => deps.logWarn(message),
    logError: (message: string, error: unknown) => deps.logError(message, error),
  });
}
