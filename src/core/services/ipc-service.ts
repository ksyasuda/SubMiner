import { BrowserWindow, ipcMain, IpcMainEvent } from "electron";

export interface IpcServiceDeps {
  getInvisibleWindow: () => BrowserWindow | null;
  isVisibleOverlayVisible: () => boolean;
  setInvisibleIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => void;
  onOverlayModalClosed: (modal: string) => void;
  openYomitanSettings: () => void;
  quitApp: () => void;
  toggleDevTools: () => void;
  getVisibleOverlayVisibility: () => boolean;
  toggleVisibleOverlay: () => void;
  getInvisibleOverlayVisibility: () => boolean;
  tokenizeCurrentSubtitle: () => Promise<unknown>;
  getCurrentSubtitleAss: () => string;
  getMpvSubtitleRenderMetrics: () => unknown;
  getSubtitlePosition: () => unknown;
  getSubtitleStyle: () => unknown;
  saveSubtitlePosition: (position: unknown) => void;
  getMecabStatus: () => { available: boolean; enabled: boolean; path: string | null };
  setMecabEnabled: (enabled: boolean) => void;
  handleMpvCommand: (command: Array<string | number>) => void;
  getKeybindings: () => unknown;
  getSecondarySubMode: () => unknown;
  getCurrentSecondarySub: () => string;
  runSubsyncManual: (request: unknown) => Promise<unknown>;
  getAnkiConnectStatus: () => boolean;
  getRuntimeOptions: () => unknown;
  setRuntimeOption: (id: string, value: unknown) => unknown;
  cycleRuntimeOption: (id: string, direction: 1 | -1) => unknown;
}

export function registerIpcHandlersService(deps: IpcServiceDeps): void {
  ipcMain.on(
    "set-ignore-mouse-events",
    (
      event: IpcMainEvent,
      ignore: boolean,
      options: { forward?: boolean } = {},
    ) => {
      const senderWindow = BrowserWindow.fromWebContents(event.sender);
      if (senderWindow && !senderWindow.isDestroyed()) {
        const invisibleWindow = deps.getInvisibleWindow();
        if (
          senderWindow === invisibleWindow &&
          deps.isVisibleOverlayVisible() &&
          invisibleWindow &&
          !invisibleWindow.isDestroyed()
        ) {
          deps.setInvisibleIgnoreMouseEvents(true, { forward: true });
        } else {
          senderWindow.setIgnoreMouseEvents(ignore, options);
        }
      }
    },
  );

  ipcMain.on("overlay:modal-closed", (_event: IpcMainEvent, modal: string) => {
    deps.onOverlayModalClosed(modal);
  });

  ipcMain.on("open-yomitan-settings", () => {
    deps.openYomitanSettings();
  });

  ipcMain.on("quit-app", () => {
    deps.quitApp();
  });

  ipcMain.on("toggle-dev-tools", () => {
    deps.toggleDevTools();
  });

  ipcMain.handle("get-overlay-visibility", () => {
    return deps.getVisibleOverlayVisibility();
  });

  ipcMain.on("toggle-overlay", () => {
    deps.toggleVisibleOverlay();
  });

  ipcMain.handle("get-visible-overlay-visibility", () => {
    return deps.getVisibleOverlayVisibility();
  });

  ipcMain.handle("get-invisible-overlay-visibility", () => {
    return deps.getInvisibleOverlayVisibility();
  });

  ipcMain.handle("get-current-subtitle", async () => {
    return await deps.tokenizeCurrentSubtitle();
  });

  ipcMain.handle("get-current-subtitle-ass", () => {
    return deps.getCurrentSubtitleAss();
  });

  ipcMain.handle("get-mpv-subtitle-render-metrics", () => {
    return deps.getMpvSubtitleRenderMetrics();
  });

  ipcMain.handle("get-subtitle-position", () => {
    return deps.getSubtitlePosition();
  });

  ipcMain.handle("get-subtitle-style", () => {
    return deps.getSubtitleStyle();
  });

  ipcMain.on("save-subtitle-position", (_event: IpcMainEvent, position: unknown) => {
    deps.saveSubtitlePosition(position);
  });

  ipcMain.handle("get-mecab-status", () => {
    return deps.getMecabStatus();
  });

  ipcMain.on("set-mecab-enabled", (_event: IpcMainEvent, enabled: boolean) => {
    deps.setMecabEnabled(enabled);
  });

  ipcMain.on("mpv-command", (_event: IpcMainEvent, command: (string | number)[]) => {
    deps.handleMpvCommand(command);
  });

  ipcMain.handle("get-keybindings", () => {
    return deps.getKeybindings();
  });

  ipcMain.handle("get-secondary-sub-mode", () => {
    return deps.getSecondarySubMode();
  });

  ipcMain.handle("get-current-secondary-sub", () => {
    return deps.getCurrentSecondarySub();
  });

  ipcMain.handle("subsync:run-manual", async (_event, request: unknown) => {
    return await deps.runSubsyncManual(request);
  });

  ipcMain.handle("get-anki-connect-status", () => {
    return deps.getAnkiConnectStatus();
  });

  ipcMain.handle("runtime-options:get", () => {
    return deps.getRuntimeOptions();
  });

  ipcMain.handle("runtime-options:set", (_event, id: string, value: unknown) => {
    return deps.setRuntimeOption(id, value);
  });

  ipcMain.handle("runtime-options:cycle", (_event, id: string, direction: 1 | -1) => {
    return deps.cycleRuntimeOption(id, direction);
  });
}
