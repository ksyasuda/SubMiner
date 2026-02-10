import { BrowserWindow } from "electron";

export function sendToVisibleOverlayService(options: {
  mainWindow: BrowserWindow | null;
  visibleOverlayVisible: boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  channel: string;
  payload?: unknown;
  restoreOnModalClose?: string;
  addRestoreFlag: (modal: string) => void;
}): boolean {
  if (!options.mainWindow || options.mainWindow.isDestroyed()) return false;
  const wasVisible = options.visibleOverlayVisible;
  if (!options.visibleOverlayVisible) {
    options.setVisibleOverlayVisible(true);
  }
  if (!wasVisible && options.restoreOnModalClose) {
    options.addRestoreFlag(options.restoreOnModalClose);
  }
  if (options.payload === undefined) {
    options.mainWindow.webContents.send(options.channel);
  } else {
    options.mainWindow.webContents.send(options.channel, options.payload);
  }
  return true;
}
