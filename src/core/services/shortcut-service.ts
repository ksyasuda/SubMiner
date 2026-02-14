import { BrowserWindow, globalShortcut } from "electron";

export interface GlobalShortcutConfig {
  toggleVisibleOverlayGlobal: string | null | undefined;
  toggleInvisibleOverlayGlobal: string | null | undefined;
  openJimaku?: string | null | undefined;
}

export interface RegisterGlobalShortcutsServiceOptions {
  shortcuts: GlobalShortcutConfig;
  onToggleVisibleOverlay: () => void;
  onToggleInvisibleOverlay: () => void;
  onOpenYomitanSettings: () => void;
  onOpenJimaku?: () => void;
  isDev: boolean;
  getMainWindow: () => BrowserWindow | null;
}

export function registerGlobalShortcutsService(
  options: RegisterGlobalShortcutsServiceOptions,
): void {
  const visibleShortcut = options.shortcuts.toggleVisibleOverlayGlobal;
  const invisibleShortcut = options.shortcuts.toggleInvisibleOverlayGlobal;
  const normalizedVisible = visibleShortcut?.replace(/\s+/g, "").toLowerCase();
  const normalizedInvisible = invisibleShortcut
    ?.replace(/\s+/g, "")
    .toLowerCase();
  const normalizedJimaku = options.shortcuts.openJimaku
    ?.replace(/\s+/g, "")
    .toLowerCase();
  const normalizedSettings = "alt+shift+y";

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

  if (options.shortcuts.openJimaku && options.onOpenJimaku) {
    if (
      normalizedJimaku &&
      (normalizedJimaku === normalizedVisible ||
        normalizedJimaku === normalizedInvisible ||
        normalizedJimaku === normalizedSettings)
    ) {
      console.warn(
        "Skipped registering openJimaku because it collides with another global shortcut",
      );
    } else {
      const openJimakuRegistered = globalShortcut.register(
        options.shortcuts.openJimaku,
        () => {
          options.onOpenJimaku?.();
        },
      );
      if (!openJimakuRegistered) {
        console.warn(
          `Failed to register global shortcut openJimaku: ${options.shortcuts.openJimaku}`,
        );
      }
    }
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
