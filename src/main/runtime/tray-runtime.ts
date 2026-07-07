import { i18n } from '../../i18n/index.js';

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
      ? '\u2713 ' + i18n.t('tray.jellyfinDiscovery')
      : i18n.t('tray.jellyfinDiscovery');

  return [
    {
      label: i18n.t('tray.openHelp'),
      click: handlers.openSessionHelp,
    },
    ...(handlers.showTexthookerPage
      ? [
          {
            label: i18n.t('tray.openTexthooker'),
            click: handlers.openTexthookerInBrowser,
          },
        ]
      : []),
    ...(handlers.showFirstRunSetup
      ? [
          {
            label: i18n.t('tray.completeSetup'),
            click: handlers.openFirstRunSetup,
          },
        ]
      : []),
    ...(handlers.showWindowsMpvLauncherSetup
      ? [
          {
            label: i18n.t('tray.openSetup'),
            click: handlers.openWindowsMpvLauncherSetup,
          },
        ]
      : []),
    {
      label: i18n.t('tray.openYomitanSettings'),
      click: handlers.openYomitanSettings,
    },
    {
      label: i18n.t('tray.openSettings'),
      click: handlers.openConfigSettings,
    },
    {
      label: i18n.t('tray.exportLogs'),
      click: handlers.exportLogs,
    },
    {
      label: i18n.t('tray.configureJellyfin'),
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
      label: i18n.t('tray.configureAnilist'),
      click: handlers.openAnilistSetup,
    },
    {
      label: i18n.t('tray.checkUpdates'),
      click: handlers.checkForUpdates,
    },
    { type: 'separator' },
    {
      label: i18n.t('tray.quit'),
      click: handlers.quitApp,
    },
  ];
}
