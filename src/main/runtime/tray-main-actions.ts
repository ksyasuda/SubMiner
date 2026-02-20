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
    openOverlay: () => void;
    openYomitanSettings: () => void;
    openRuntimeOptions: () => void;
    openJellyfinSetup: () => void;
    openAnilistSetup: () => void;
    quitApp: () => void;
  }) => TMenuItem[];
  initializeOverlayRuntime: () => void;
  isOverlayRuntimeInitialized: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  openYomitanSettings: () => void;
  openRuntimeOptionsPalette: () => void;
  openJellyfinSetupWindow: () => void;
  openAnilistSetupWindow: () => void;
  quitApp: () => void;
}) {
  return (): TMenuItem[] => {
    return deps.buildTrayMenuTemplateRuntime({
      openOverlay: () => {
        if (!deps.isOverlayRuntimeInitialized()) {
          deps.initializeOverlayRuntime();
        }
        deps.setVisibleOverlayVisible(true);
      },
      openYomitanSettings: () => {
        deps.openYomitanSettings();
      },
      openRuntimeOptions: () => {
        if (!deps.isOverlayRuntimeInitialized()) {
          deps.initializeOverlayRuntime();
        }
        deps.openRuntimeOptionsPalette();
      },
      openJellyfinSetup: () => {
        deps.openJellyfinSetupWindow();
      },
      openAnilistSetup: () => {
        deps.openAnilistSetupWindow();
      },
      quitApp: () => {
        deps.quitApp();
      },
    });
  };
}
