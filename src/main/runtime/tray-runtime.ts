export function resolveTrayIconPathRuntime(deps: {
  platform: string;
  resourcesPath: string;
  appPath: string;
  dirname: string;
  joinPath: (...parts: string[]) => string;
  fileExists: (path: string) => boolean;
}): string | null {
  const iconNames =
    deps.platform === 'darwin'
      ? ['SubMinerTemplate.png', 'SubMinerTemplate@2x.png', 'SubMiner.png']
      : ['SubMiner.png'];

  const baseDirs = [
    deps.joinPath(deps.resourcesPath, 'assets'),
    deps.joinPath(deps.appPath, 'assets'),
    deps.joinPath(deps.dirname, '..', 'assets'),
    deps.joinPath(deps.dirname, '..', '..', 'assets'),
  ];

  for (const baseDir of baseDirs) {
    for (const iconName of iconNames) {
      const candidate = deps.joinPath(baseDir, iconName);
      if (deps.fileExists(candidate)) {
        return candidate;
      }
    }
  }
  return null;
}

export type TrayMenuActionHandlers = {
  openOverlay: () => void;
  openFirstRunSetup: () => void;
  showFirstRunSetup: boolean;
  openWindowsMpvLauncherSetup: () => void;
  showWindowsMpvLauncherSetup: boolean;
  openYomitanSettings: () => void;
  openRuntimeOptions: () => void;
  openJellyfinSetup: () => void;
  openAnilistSetup: () => void;
  quitApp: () => void;
};

export function buildTrayMenuTemplateRuntime(handlers: TrayMenuActionHandlers): Array<{
  label?: string;
  type?: 'separator';
  click?: () => void;
}> {
  return [
    {
      label: 'Open Overlay',
      click: handlers.openOverlay,
    },
    ...(handlers.showFirstRunSetup
      ? [
          {
            label: 'Complete Setup',
            click: handlers.openFirstRunSetup,
          },
        ]
      : []),
    ...(handlers.showWindowsMpvLauncherSetup
      ? [
          {
            label: 'Manage Windows mpv launcher',
            click: handlers.openWindowsMpvLauncherSetup,
          },
        ]
      : []),
    {
      label: 'Open Yomitan Settings',
      click: handlers.openYomitanSettings,
    },
    {
      label: 'Open Runtime Options',
      click: handlers.openRuntimeOptions,
    },
    {
      label: 'Configure Jellyfin',
      click: handlers.openJellyfinSetup,
    },
    {
      label: 'Configure AniList',
      click: handlers.openAnilistSetup,
    },
    { type: 'separator' },
    {
      label: 'Quit',
      click: handlers.quitApp,
    },
  ];
}
