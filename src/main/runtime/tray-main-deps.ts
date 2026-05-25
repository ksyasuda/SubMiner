export function createBuildResolveTrayIconPathMainDepsHandler(deps: {
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
  return () => ({
    resolveTrayIconPathRuntime: deps.resolveTrayIconPathRuntime,
    platform: deps.platform,
    resourcesPath: deps.resourcesPath,
    appPath: deps.appPath,
    dirname: deps.dirname,
    joinPath: deps.joinPath,
    fileExists: deps.fileExists,
  });
}

export function createBuildTrayMenuTemplateMainDepsHandler<TMenuItem>(deps: {
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
  openJellyfinSetupWindow: () => void;
  isJellyfinConfigured: () => boolean;
  isJellyfinDiscoveryActive: () => boolean;
  toggleJellyfinDiscovery: (checked: boolean) => void | Promise<void>;
  platform?: string;
  openAnilistSetupWindow: () => void;
  checkForUpdates: () => void;
  quitApp: () => void;
}) {
  return () => ({
    buildTrayMenuTemplateRuntime: deps.buildTrayMenuTemplateRuntime,
    platform: deps.platform,
    initializeOverlayRuntime: deps.initializeOverlayRuntime,
    isOverlayRuntimeInitialized: deps.isOverlayRuntimeInitialized,
    openSessionHelpModal: deps.openSessionHelpModal,
    openTexthookerInBrowser: deps.openTexthookerInBrowser,
    showTexthookerPage: deps.showTexthookerPage,
    showFirstRunSetup: deps.showFirstRunSetup,
    openFirstRunSetupWindow: deps.openFirstRunSetupWindow,
    showWindowsMpvLauncherSetup: deps.showWindowsMpvLauncherSetup,
    openYomitanSettings: deps.openYomitanSettings,
    openConfigSettingsWindow: deps.openConfigSettingsWindow,
    openJellyfinSetupWindow: deps.openJellyfinSetupWindow,
    isJellyfinConfigured: deps.isJellyfinConfigured,
    isJellyfinDiscoveryActive: deps.isJellyfinDiscoveryActive,
    toggleJellyfinDiscovery: deps.toggleJellyfinDiscovery,
    openAnilistSetupWindow: deps.openAnilistSetupWindow,
    checkForUpdates: deps.checkForUpdates,
    quitApp: deps.quitApp,
  });
}
