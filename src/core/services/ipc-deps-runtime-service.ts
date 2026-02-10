import { IpcServiceDeps } from "./ipc-service";

interface WindowLike {
  isDestroyed: () => boolean;
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
  getSecondarySubMode: () => unknown;
  getMpvClient: () => MpvClientLike | null;
  runSubsyncManual: (request: unknown) => Promise<unknown>;
  getAnkiConnectStatus: () => boolean;
  getRuntimeOptions: () => unknown;
  setRuntimeOption: (id: string, value: unknown) => unknown;
  cycleRuntimeOption: (id: string, direction: 1 | -1) => unknown;
}

export function createIpcDepsRuntimeService(
  options: IpcDepsRuntimeOptions,
): IpcServiceDeps {
  return {
    getInvisibleWindow: () => options.getInvisibleWindow() as never,
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
    getSecondarySubMode: options.getSecondarySubMode,
    getCurrentSecondarySub: () =>
      options.getMpvClient()?.currentSecondarySubText || "",
    runSubsyncManual: options.runSubsyncManual,
    getAnkiConnectStatus: options.getAnkiConnectStatus,
    getRuntimeOptions: options.getRuntimeOptions,
    setRuntimeOption: options.setRuntimeOption,
    cycleRuntimeOption: options.cycleRuntimeOption,
  };
}
