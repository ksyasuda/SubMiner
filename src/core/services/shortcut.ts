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
