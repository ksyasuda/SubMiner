export function createBuildEnsureTrayMainDepsHandler(deps: {
  getTray: () => unknown | null;
  setTray: (tray: unknown | null) => void;
  buildTrayMenu: () => unknown;
  resolveTrayIconPath: () => string | null;
  createImageFromPath: (iconPath: string) => unknown;
  createEmptyImage: () => unknown;
  createTray: (icon: unknown) => unknown;
  trayTooltip: string;
  platform: string;
  logWarn: (message: string) => void;
  initializeOverlayRuntime: () => void;
  isOverlayRuntimeInitialized: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
}) {
  return () => ({
    getTray: () => deps.getTray() as never,
    setTray: (tray: unknown | null) => deps.setTray(tray),
    buildTrayMenu: () => deps.buildTrayMenu() as never,
    resolveTrayIconPath: () => deps.resolveTrayIconPath(),
    createImageFromPath: (iconPath: string) => deps.createImageFromPath(iconPath) as never,
    createEmptyImage: () => deps.createEmptyImage() as never,
    createTray: (icon: unknown) => deps.createTray(icon) as never,
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

export function createBuildDestroyTrayMainDepsHandler(deps: {
  getTray: () => unknown | null;
  setTray: (tray: unknown | null) => void;
}) {
  return () => ({
    getTray: () => deps.getTray() as never,
    setTray: (tray: unknown | null) => deps.setTray(tray),
  });
}

export function createBuildInitializeOverlayRuntimeBootstrapMainDepsHandler(deps: {
  isOverlayRuntimeInitialized: () => boolean;
  initializeOverlayRuntimeCore: (options: unknown) => { invisibleOverlayVisible: boolean };
  buildOptions: () => unknown;
  setInvisibleOverlayVisible: (visible: boolean) => void;
  setOverlayRuntimeInitialized: (initialized: boolean) => void;
  startBackgroundWarmups: () => void;
}) {
  return () => ({
    isOverlayRuntimeInitialized: () => deps.isOverlayRuntimeInitialized(),
    initializeOverlayRuntimeCore: (options: unknown) => deps.initializeOverlayRuntimeCore(options),
    buildOptions: () => deps.buildOptions() as never,
    setInvisibleOverlayVisible: (visible: boolean) => deps.setInvisibleOverlayVisible(visible),
    setOverlayRuntimeInitialized: (initialized: boolean) =>
      deps.setOverlayRuntimeInitialized(initialized),
    startBackgroundWarmups: () => deps.startBackgroundWarmups(),
  });
}

export function createBuildOpenYomitanSettingsMainDepsHandler(deps: {
  ensureYomitanExtensionLoaded: () => Promise<unknown | null>;
  openYomitanSettingsWindow: (params: {
    yomitanExt: unknown;
    getExistingWindow: () => unknown | null;
    setWindow: (window: unknown | null) => void;
  }) => void;
  getExistingWindow: () => unknown | null;
  setWindow: (window: unknown | null) => void;
  logWarn: (message: string) => void;
  logError: (message: string, error: unknown) => void;
}) {
  return () => ({
    ensureYomitanExtensionLoaded: () => deps.ensureYomitanExtensionLoaded(),
    openYomitanSettingsWindow: (params: {
      yomitanExt: unknown;
      getExistingWindow: () => unknown | null;
      setWindow: (window: unknown | null) => void;
    }) => deps.openYomitanSettingsWindow(params),
    getExistingWindow: () => deps.getExistingWindow(),
    setWindow: (window: unknown | null) => deps.setWindow(window),
    logWarn: (message: string) => deps.logWarn(message),
    logError: (message: string, error: unknown) => deps.logError(message, error),
  });
}
