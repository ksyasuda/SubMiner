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
  openSyncUi: () => void;
  exportLogs: () => void;
  openJellyfinSetup: () => void;
  showJellyfinDiscovery: boolean;
  jellyfinDiscoveryActive: boolean;
  toggleJellyfinDiscovery: (checked: boolean) => void;
  openAnilistSetup: () => void;
  checkForUpdates: () => void;
  quitApp: () => void;
};

type TrayMenuClickItem = {
  checked?: boolean;
};

export function buildTrayMenuTemplateRuntime(handlers: TrayMenuActionHandlers): Array<{
  label?: string;
  type?: 'separator' | 'checkbox';
  checked?: boolean;
  enabled?: boolean;
  click?: (menuItem?: TrayMenuClickItem) => void;
}> {
  const jellyfinDiscoveryLabel =
    handlers.platform === 'linux' && handlers.jellyfinDiscoveryActive
      ? '✓ Jellyfin Discovery'
      : 'Jellyfin Discovery';

  return [
    {
      label: 'Open Help',
      click: handlers.openSessionHelp,
    },
    ...(handlers.showTexthookerPage
      ? [
          {
            label: 'Open Texthooker',
            click: handlers.openTexthookerInBrowser,
          },
        ]
      : []),
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
            label: 'Open SubMiner Setup',
            click: handlers.openWindowsMpvLauncherSetup,
          },
        ]
      : []),
    {
      label: 'Open Yomitan Settings',
      click: handlers.openYomitanSettings,
    },
    {
      label: 'Open SubMiner Settings',
      click: handlers.openConfigSettings,
    },
    {
      label: 'Sync Stats && History',
      click: handlers.openSyncUi,
    },
    {
      label: 'Export Logs',
      click: handlers.exportLogs,
    },
    {
      label: 'Configure Jellyfin',
      click: handlers.openJellyfinSetup,
    },
    ...(handlers.showJellyfinDiscovery
      ? [
          {
            label: jellyfinDiscoveryLabel,
            type: 'checkbox' as const,
            checked: handlers.jellyfinDiscoveryActive,
            enabled: true,
            click: (menuItem?: TrayMenuClickItem) => {
              const checked =
                typeof menuItem?.checked === 'boolean'
                  ? menuItem.checked
                  : !handlers.jellyfinDiscoveryActive;
              handlers.toggleJellyfinDiscovery(checked);
            },
          },
        ]
      : []),
    {
      label: 'Configure AniList',
      click: handlers.openAnilistSetup,
    },
    {
      label: 'Check for Updates',
      click: handlers.checkForUpdates,
    },
    { type: 'separator' },
    {
      label: 'Quit',
      click: handlers.quitApp,
    },
  ];
}
