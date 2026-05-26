export function createResolveTrayIconPathHandler(deps: {
  resolveTrayIconPathRuntime: (options: {
    platform: string;
    resourcesPath: string;
    appPath: string;
    dirname: string;
    joinPath: (...parts: string[]) => string;
    fileExists: (path: string) => boolean;
  }) => string | null;
  platform: string;
  resourcesPath: string;
  appPath: string;
  dirname: string;
  joinPath: (...parts: string[]) => string;
  fileExists: (path: string) => boolean;
}) {
  return (): string | null => {
    return deps.resolveTrayIconPathRuntime({
      platform: deps.platform,
      resourcesPath: deps.resourcesPath,
      appPath: deps.appPath,
      dirname: deps.dirname,
      joinPath: deps.joinPath,
      fileExists: deps.fileExists,
    });
  };
}

export function shouldShowTexthookerTrayEntry(config: {
  websocket?: { enabled?: boolean | 'auto' };
  annotationWebsocket?: { enabled?: boolean };
}): boolean {
  const websocketEnabled = config.websocket?.enabled ?? false;
  const annotationWebsocketEnabled = config.annotationWebsocket?.enabled ?? false;
  return websocketEnabled !== false || annotationWebsocketEnabled !== false;
}

export function createBuildTrayMenuTemplateHandler<TMenuItem>(deps: {
  buildTrayMenuTemplateRuntime: (handlers: {
    platform?: string;
    openSessionHelp: () => void;
    openTexthookerInBrowser: () => void;
    showTexthookerPage: boolean;
    openFirstRunSetup: () => void;
    showFirstRunSetup: boolean;
    openWindowsMpvLauncherSetup: () => void;
    showWindowsMpvLauncherSetup: boolean;
    openYomitanSettings: () => void;
    openConfigSettings: () => void;
    exportLogs: () => void;
    openJellyfinSetup: () => void;
    showJellyfinDiscovery: boolean;
    jellyfinDiscoveryActive: boolean;
    toggleJellyfinDiscovery: (checked: boolean) => void;
    openAnilistSetup: () => void;
    checkForUpdates: () => void;
    quitApp: () => void;
  }) => TMenuItem[];
  initializeOverlayRuntime: () => void;
  isOverlayRuntimeInitialized: () => boolean;
  openSessionHelpModal: () => void;
  openTexthookerInBrowser: () => void;
  showTexthookerPage: () => boolean;
  showFirstRunSetup: () => boolean;
  openFirstRunSetupWindow: (force?: boolean) => void;
  showWindowsMpvLauncherSetup: () => boolean;
  openYomitanSettings: () => void;
  openConfigSettingsWindow: () => void;
  exportLogs: () => void;
  openJellyfinSetupWindow: () => void;
  isJellyfinConfigured: () => boolean;
  isJellyfinDiscoveryActive: () => boolean;
  toggleJellyfinDiscovery: (checked: boolean) => void | Promise<void>;
  platform?: string;
  openAnilistSetupWindow: () => void;
  checkForUpdates: () => void;
  quitApp: () => void;
}) {
  return (): TMenuItem[] => {
    return deps.buildTrayMenuTemplateRuntime({
      platform: deps.platform,
      openSessionHelp: () => {
        if (!deps.isOverlayRuntimeInitialized()) {
          deps.initializeOverlayRuntime();
        }
        deps.openSessionHelpModal();
      },
      openTexthookerInBrowser: () => {
        deps.openTexthookerInBrowser();
      },
      showTexthookerPage: deps.showTexthookerPage(),
      openFirstRunSetup: () => {
        deps.openFirstRunSetupWindow();
      },
      showFirstRunSetup: deps.showFirstRunSetup(),
      openWindowsMpvLauncherSetup: () => {
        deps.openFirstRunSetupWindow(true);
      },
      showWindowsMpvLauncherSetup: deps.showWindowsMpvLauncherSetup(),
      openYomitanSettings: () => {
        deps.openYomitanSettings();
      },
      openConfigSettings: () => {
        deps.openConfigSettingsWindow();
      },
      exportLogs: () => {
        deps.exportLogs();
      },
      openJellyfinSetup: () => {
        deps.openJellyfinSetupWindow();
      },
      showJellyfinDiscovery: deps.isJellyfinConfigured(),
      jellyfinDiscoveryActive: deps.isJellyfinDiscoveryActive(),
      toggleJellyfinDiscovery: (checked) => {
        void deps.toggleJellyfinDiscovery(checked);
      },
      openAnilistSetup: () => {
        deps.openAnilistSetupWindow();
      },
      checkForUpdates: () => {
        deps.checkForUpdates();
      },
      quitApp: () => {
        deps.quitApp();
      },
    });
  };
}
