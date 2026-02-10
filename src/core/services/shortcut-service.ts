import { BrowserWindow, globalShortcut } from "electron";

export interface GlobalShortcutConfig {
  toggleVisibleOverlayGlobal: string | null | undefined;
  toggleInvisibleOverlayGlobal: string | null | undefined;
}

export function registerGlobalShortcutsService(options: {
  shortcuts: GlobalShortcutConfig;
  onToggleVisibleOverlay: () => void;
  onToggleInvisibleOverlay: () => void;
  onOpenYomitanSettings: () => void;
  isDev: boolean;
  getMainWindow: () => BrowserWindow | null;
}): void {
  const visibleShortcut = options.shortcuts.toggleVisibleOverlayGlobal;
  const invisibleShortcut = options.shortcuts.toggleInvisibleOverlayGlobal;
  const normalizedVisible = visibleShortcut?.replace(/\s+/g, "").toLowerCase();
  const normalizedInvisible = invisibleShortcut
    ?.replace(/\s+/g, "")
    .toLowerCase();

  if (visibleShortcut) {
    const toggleVisibleRegistered = globalShortcut.register(
      visibleShortcut,
      () => {
        options.onToggleVisibleOverlay();
      },
    );
    if (!toggleVisibleRegistered) {
      console.warn(
        `Failed to register global shortcut toggleVisibleOverlayGlobal: ${visibleShortcut}`,
      );
    }
  }

  if (
    invisibleShortcut &&
    normalizedInvisible &&
    normalizedInvisible !== normalizedVisible
  ) {
    const toggleInvisibleRegistered = globalShortcut.register(
      invisibleShortcut,
      () => {
        options.onToggleInvisibleOverlay();
      },
    );
    if (!toggleInvisibleRegistered) {
      console.warn(
        `Failed to register global shortcut toggleInvisibleOverlayGlobal: ${invisibleShortcut}`,
      );
    }
  } else if (
    invisibleShortcut &&
    normalizedInvisible &&
    normalizedInvisible === normalizedVisible
  ) {
    console.warn(
      "Skipped registering toggleInvisibleOverlayGlobal because it collides with toggleVisibleOverlayGlobal",
    );
  }

  const settingsRegistered = globalShortcut.register("Alt+Shift+Y", () => {
    options.onOpenYomitanSettings();
  });
  if (!settingsRegistered) {
    console.warn("Failed to register global shortcut: Alt+Shift+Y");
  }

  if (options.isDev) {
    const devtoolsRegistered = globalShortcut.register("F12", () => {
      const mainWindow = options.getMainWindow();
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.toggleDevTools();
      }
    });
    if (!devtoolsRegistered) {
      console.warn("Failed to register global shortcut: F12");
    }
  }
}
