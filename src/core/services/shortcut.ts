import electron from 'electron';
import type { BrowserWindow } from 'electron';
import { createLogger } from '../../logger';

const { globalShortcut } = electron;
const logger = createLogger('main:shortcut');

export interface GlobalShortcutConfig {
  toggleVisibleOverlayGlobal: string | null | undefined;
  openJimaku?: string | null | undefined;
}

export interface RegisterGlobalShortcutsServiceOptions {
  shortcuts: GlobalShortcutConfig;
  onToggleVisibleOverlay: () => void;
  onOpenYomitanSettings: () => void;
  onOpenJimaku?: () => void;
  isDev: boolean;
  getMainWindow: () => BrowserWindow | null;
}

export function registerGlobalShortcuts(options: RegisterGlobalShortcutsServiceOptions): void {
  const visibleShortcut = options.shortcuts.toggleVisibleOverlayGlobal;
  const normalizedVisible = visibleShortcut?.replace(/\s+/g, '').toLowerCase();
  const normalizedJimaku = options.shortcuts.openJimaku?.replace(/\s+/g, '').toLowerCase();
  const normalizedSettings = 'alt+shift+y';

  if (visibleShortcut) {
    const toggleVisibleRegistered = globalShortcut.register(visibleShortcut, () => {
      options.onToggleVisibleOverlay();
    });
    if (!toggleVisibleRegistered) {
      logger.warn(
        `Failed to register global shortcut toggleVisibleOverlayGlobal: ${visibleShortcut}`,
      );
    }
  }

  if (options.shortcuts.openJimaku && options.onOpenJimaku) {
    if (
      normalizedJimaku &&
      (normalizedJimaku === normalizedVisible || normalizedJimaku === normalizedSettings)
    ) {
      logger.warn(
        'Skipped registering openJimaku because it collides with another global shortcut',
      );
    } else {
      const openJimakuRegistered = globalShortcut.register(options.shortcuts.openJimaku, () => {
        options.onOpenJimaku?.();
      });
      if (!openJimakuRegistered) {
        logger.warn(
          `Failed to register global shortcut openJimaku: ${options.shortcuts.openJimaku}`,
        );
      }
    }
  }

  const settingsRegistered = globalShortcut.register('Alt+Shift+Y', () => {
    options.onOpenYomitanSettings();
  });
  if (!settingsRegistered) {
    logger.warn('Failed to register global shortcut: Alt+Shift+Y');
  }

  if (options.isDev) {
    const devtoolsRegistered = globalShortcut.register('F12', () => {
      const mainWindow = options.getMainWindow();
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.toggleDevTools();
      }
    });
    if (!devtoolsRegistered) {
      logger.warn('Failed to register global shortcut: F12');
    }
  }
}
