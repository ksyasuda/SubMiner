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
    openOverlay: () => void;
    openFirstRunSetup: () => void;
    showFirstRunSetup: boolean;
    openYomitanSettings: () => void;
    openRuntimeOptions: () => void;
    openJellyfinSetup: () => void;
    openAnilistSetup: () => void;
    quitApp: () => void;
  }) => TMenuItem[];
  initializeOverlayRuntime: () => void;
  isOverlayRuntimeInitialized: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  showFirstRunSetup: () => boolean;
  openFirstRunSetupWindow: () => void;
  openYomitanSettings: () => void;
  openRuntimeOptionsPalette: () => void;
  openJellyfinSetupWindow: () => void;
  openAnilistSetupWindow: () => void;
  quitApp: () => void;
}) {
  return () => ({
    buildTrayMenuTemplateRuntime: deps.buildTrayMenuTemplateRuntime,
    initializeOverlayRuntime: deps.initializeOverlayRuntime,
    isOverlayRuntimeInitialized: deps.isOverlayRuntimeInitialized,
    setVisibleOverlayVisible: deps.setVisibleOverlayVisible,
    showFirstRunSetup: deps.showFirstRunSetup,
    openFirstRunSetupWindow: deps.openFirstRunSetupWindow,
    openYomitanSettings: deps.openYomitanSettings,
    openRuntimeOptionsPalette: deps.openRuntimeOptionsPalette,
    openJellyfinSetupWindow: deps.openJellyfinSetupWindow,
    openAnilistSetupWindow: deps.openAnilistSetupWindow,
    quitApp: deps.quitApp,
  });
}
