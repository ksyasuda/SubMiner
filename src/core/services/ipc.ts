import { BrowserWindow, ipcMain, IpcMainEvent } from "electron";

export interface IpcServiceDeps {
  getInvisibleWindow: () => WindowLike | null;
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
  getConfiguredShortcuts: () => unknown;
  getSecondarySubMode: () => unknown;
  getCurrentSecondarySub: () => string;
  focusMainWindow: () => void;
  runSubsyncManual: (request: unknown) => Promise<unknown>;
  getAnkiConnectStatus: () => boolean;
  getRuntimeOptions: () => unknown;
  setRuntimeOption: (id: string, value: unknown) => unknown;
  cycleRuntimeOption: (id: string, direction: 1 | -1) => unknown;
  reportOverlayContentBounds: (payload: unknown) => void;
  getAnilistStatus: () => unknown;
  clearAnilistToken: () => void;
  openAnilistSetup: () => void;
  getAnilistQueueStatus: () => unknown;
  retryAnilistQueueNow: () => Promise<{ ok: boolean; message: string }>;
}

interface WindowLike {
  isDestroyed: () => boolean;
  focus: () => void;
  setIgnoreMouseEvents: (
    ignore: boolean,
    options?: { forward?: boolean },
  ) => void;
  webContents: {
    toggleDevTools: () => void;
  };
}

interface MecabTokenizerLike {
  getStatus: () => { available: boolean; enabled: boolean; path: string | null };
  setEnabled: (enabled: boolean) => void;
}

interface MpvClientLike {
  currentSecondarySubText?: string;
}

export interface IpcDepsRuntimeOptions {
  getInvisibleWindow: () => WindowLike | null;
  getMainWindow: () => WindowLike | null;
  getVisibleOverlayVisibility: () => boolean;
  getInvisibleOverlayVisibility: () => boolean;
  onOverlayModalClosed: (modal: string) => void;
  openYomitanSettings: () => void;
  quitApp: () => void;
  toggleVisibleOverlay: () => void;
  tokenizeCurrentSubtitle: () => Promise<unknown>;
  getCurrentSubtitleAss: () => string;
  getMpvSubtitleRenderMetrics: () => unknown;
  getSubtitlePosition: () => unknown;
  getSubtitleStyle: () => unknown;
  saveSubtitlePosition: (position: unknown) => void;
  getMecabTokenizer: () => MecabTokenizerLike | null;
  handleMpvCommand: (command: Array<string | number>) => void;
  getKeybindings: () => unknown;
  getConfiguredShortcuts: () => unknown;
  getSecondarySubMode: () => unknown;
  getMpvClient: () => MpvClientLike | null;
  focusMainWindow: () => void;
  runSubsyncManual: (request: unknown) => Promise<unknown>;
  getAnkiConnectStatus: () => boolean;
  getRuntimeOptions: () => unknown;
  setRuntimeOption: (id: string, value: unknown) => unknown;
  cycleRuntimeOption: (id: string, direction: 1 | -1) => unknown;
  reportOverlayContentBounds: (payload: unknown) => void;
  getAnilistStatus: () => unknown;
  clearAnilistToken: () => void;
  openAnilistSetup: () => void;
  getAnilistQueueStatus: () => unknown;
  retryAnilistQueueNow: () => Promise<{ ok: boolean; message: string }>;
}

export function createIpcDepsRuntime(
  options: IpcDepsRuntimeOptions,
): IpcServiceDeps {
  return {
    getInvisibleWindow: () => options.getInvisibleWindow(),
    isVisibleOverlayVisible: options.getVisibleOverlayVisibility,
    setInvisibleIgnoreMouseEvents: (ignore, eventsOptions) => {
      const invisibleWindow = options.getInvisibleWindow();
      if (!invisibleWindow || invisibleWindow.isDestroyed()) return;
      invisibleWindow.setIgnoreMouseEvents(ignore, eventsOptions);
    },
    onOverlayModalClosed: options.onOverlayModalClosed,
    openYomitanSettings: options.openYomitanSettings,
    quitApp: options.quitApp,
    toggleDevTools: () => {
      const mainWindow = options.getMainWindow();
      if (!mainWindow || mainWindow.isDestroyed()) return;
      mainWindow.webContents.toggleDevTools();
    },
    getVisibleOverlayVisibility: options.getVisibleOverlayVisibility,
    toggleVisibleOverlay: options.toggleVisibleOverlay,
    getInvisibleOverlayVisibility: options.getInvisibleOverlayVisibility,
    tokenizeCurrentSubtitle: options.tokenizeCurrentSubtitle,
    getCurrentSubtitleAss: options.getCurrentSubtitleAss,
    getMpvSubtitleRenderMetrics: options.getMpvSubtitleRenderMetrics,
    getSubtitlePosition: options.getSubtitlePosition,
    getSubtitleStyle: options.getSubtitleStyle,
    saveSubtitlePosition: options.saveSubtitlePosition,
    getMecabStatus: () => {
      const mecabTokenizer = options.getMecabTokenizer();
      return mecabTokenizer
        ? mecabTokenizer.getStatus()
        : { available: false, enabled: false, path: null };
    },
    setMecabEnabled: (enabled) => {
      const mecabTokenizer = options.getMecabTokenizer();
      if (!mecabTokenizer) return;
      mecabTokenizer.setEnabled(enabled);
    },
    handleMpvCommand: options.handleMpvCommand,
    getKeybindings: options.getKeybindings,
    getConfiguredShortcuts: options.getConfiguredShortcuts,
    getSecondarySubMode: options.getSecondarySubMode,
    getCurrentSecondarySub: () =>
      options.getMpvClient()?.currentSecondarySubText || "",
    focusMainWindow: () => {
      const mainWindow = options.getMainWindow();
      if (!mainWindow || mainWindow.isDestroyed()) return;
      mainWindow.focus();
    },
    runSubsyncManual: options.runSubsyncManual,
    getAnkiConnectStatus: options.getAnkiConnectStatus,
    getRuntimeOptions: options.getRuntimeOptions,
    setRuntimeOption: options.setRuntimeOption,
    cycleRuntimeOption: options.cycleRuntimeOption,
    reportOverlayContentBounds: options.reportOverlayContentBounds,
    getAnilistStatus: options.getAnilistStatus,
    clearAnilistToken: options.clearAnilistToken,
    openAnilistSetup: options.openAnilistSetup,
    getAnilistQueueStatus: options.getAnilistQueueStatus,
    retryAnilistQueueNow: options.retryAnilistQueueNow,
  };
}

export function registerIpcHandlers(deps: IpcServiceDeps): void {
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

  ipcMain.handle("get-config-shortcuts", () => {
    return deps.getConfiguredShortcuts();
  });

  ipcMain.handle("get-secondary-sub-mode", () => {
    return deps.getSecondarySubMode();
  });

  ipcMain.handle("get-current-secondary-sub", () => {
    return deps.getCurrentSecondarySub();
  });

  ipcMain.handle("focus-main-window", () => {
    deps.focusMainWindow();
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

  ipcMain.on("overlay-content-bounds:report", (_event: IpcMainEvent, payload: unknown) => {
    deps.reportOverlayContentBounds(payload);
  });

  ipcMain.handle("anilist:get-status", () => {
    return deps.getAnilistStatus();
  });

  ipcMain.handle("anilist:clear-token", () => {
    deps.clearAnilistToken();
    return { ok: true };
  });

  ipcMain.handle("anilist:open-setup", () => {
    deps.openAnilistSetup();
    return { ok: true };
  });

  ipcMain.handle("anilist:get-queue-status", () => {
    return deps.getAnilistQueueStatus();
  });

  ipcMain.handle("anilist:retry-now", async () => {
    return await deps.retryAnilistQueueNow();
  });
}
