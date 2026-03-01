type TrayIconLike = {
  isEmpty: () => boolean;
  resize: (options: {
    width: number;
    height: number;
    quality?: 'best' | 'better' | 'good';
  }) => TrayIconLike;
  setTemplateImage: (enabled: boolean) => void;
};

type TrayLike = {
  setContextMenu: (menu: any) => void;
  setToolTip: (tooltip: string) => void;
  on: (event: 'click', handler: () => void) => void;
  destroy: () => void;
};

export function createEnsureTrayHandler(deps: {
  getTray: () => TrayLike | null;
  setTray: (tray: TrayLike | null) => void;
  buildTrayMenu: () => any;
  resolveTrayIconPath: () => string | null;
  createImageFromPath: (iconPath: string) => TrayIconLike;
  createEmptyImage: () => TrayIconLike;
  createTray: (icon: TrayIconLike) => TrayLike;
  trayTooltip: string;
  platform: string;
  logWarn: (message: string) => void;
  ensureOverlayVisibleFromTrayClick: () => void;
}) {
  return (): void => {
    const existingTray = deps.getTray();
    if (existingTray) {
      existingTray.setContextMenu(deps.buildTrayMenu());
      return;
    }

    const iconPath = deps.resolveTrayIconPath();
    let trayIcon = iconPath ? deps.createImageFromPath(iconPath) : deps.createEmptyImage();
    if (trayIcon.isEmpty()) {
      deps.logWarn('Tray icon asset not found; using empty icon placeholder.');
    }
    if (deps.platform === 'darwin' && !trayIcon.isEmpty()) {
      trayIcon = trayIcon.resize({ width: 18, height: 18, quality: 'best' });
      trayIcon.setTemplateImage(true);
    }
    if (deps.platform === 'linux' && !trayIcon.isEmpty()) {
      trayIcon = trayIcon.resize({ width: 20, height: 20 });
    }

    const tray = deps.createTray(trayIcon);
    tray.setToolTip(deps.trayTooltip);
    tray.setContextMenu(deps.buildTrayMenu());
    tray.on('click', () => {
      deps.ensureOverlayVisibleFromTrayClick();
    });
    deps.setTray(tray);
  };
}

export function createDestroyTrayHandler(deps: {
  getTray: () => TrayLike | null;
  setTray: (tray: TrayLike | null) => void;
}) {
  return (): void => {
    const tray = deps.getTray();
    if (!tray) return;
    tray.destroy();
    deps.setTray(null);
  };
}
