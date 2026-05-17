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

export function createBuildTrayMenuTemplateHandler<TMenuItem>(deps: {
  buildTrayMenuTemplateRuntime: (handlers: {
    openSessionHelp: () => void;
    openTexthookerInBrowser: () => void;
    showTexthookerPage: boolean;
    openFirstRunSetup: () => void;
    showFirstRunSetup: boolean;
    openWindowsMpvLauncherSetup: () => void;
    showWindowsMpvLauncherSetup: boolean;
    openYomitanSettings: () => void;
    openRuntimeOptions: () => void;
    openConfigSettings: () => void;
    openJellyfinSetup: () => void;
    showJellyfinDiscovery: boolean;
    jellyfinDiscoveryActive: boolean;
    toggleJellyfinDiscovery: () => void;
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
  openFirstRunSetupWindow: () => void;
  showWindowsMpvLauncherSetup: () => boolean;
  openYomitanSettings: () => void;
  openRuntimeOptionsPalette: () => void;
  openConfigSettingsWindow: () => void;
  openJellyfinSetupWindow: () => void;
  isJellyfinConfigured: () => boolean;
  isJellyfinDiscoveryActive: () => boolean;
  toggleJellyfinDiscovery: () => void | Promise<void>;
  openAnilistSetupWindow: () => void;
  checkForUpdates: () => void;
  quitApp: () => void;
}) {
  return (): TMenuItem[] => {
    return deps.buildTrayMenuTemplateRuntime({
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
        deps.openFirstRunSetupWindow();
      },
      showWindowsMpvLauncherSetup: deps.showWindowsMpvLauncherSetup(),
      openYomitanSettings: () => {
        deps.openYomitanSettings();
      },
      openRuntimeOptions: () => {
        if (!deps.isOverlayRuntimeInitialized()) {
          deps.initializeOverlayRuntime();
        }
        deps.openRuntimeOptionsPalette();
      },
      openConfigSettings: () => {
        deps.openConfigSettingsWindow();
      },
      openJellyfinSetup: () => {
        deps.openJellyfinSetupWindow();
      },
      showJellyfinDiscovery: deps.isJellyfinConfigured(),
      jellyfinDiscoveryActive: deps.isJellyfinDiscoveryActive(),
      toggleJellyfinDiscovery: () => {
        void deps.toggleJellyfinDiscovery();
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
