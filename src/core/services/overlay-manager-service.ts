import { BrowserWindow } from "electron";

export interface OverlayManagerService {
  getMainWindow: () => BrowserWindow | null;
  setMainWindow: (window: BrowserWindow | null) => void;
  getInvisibleWindow: () => BrowserWindow | null;
  setInvisibleWindow: (window: BrowserWindow | null) => void;
  getVisibleOverlayVisible: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  getInvisibleOverlayVisible: () => boolean;
  setInvisibleOverlayVisible: (visible: boolean) => void;
  getOverlayWindows: () => BrowserWindow[];
}

export function createOverlayManagerService(): OverlayManagerService {
  let mainWindow: BrowserWindow | null = null;
  let invisibleWindow: BrowserWindow | null = null;
  let visibleOverlayVisible = false;
  let invisibleOverlayVisible = false;

  return {
    getMainWindow: () => mainWindow,
    setMainWindow: (window) => {
      mainWindow = window;
    },
    getInvisibleWindow: () => invisibleWindow,
    setInvisibleWindow: (window) => {
      invisibleWindow = window;
    },
    getVisibleOverlayVisible: () => visibleOverlayVisible,
    setVisibleOverlayVisible: (visible) => {
      visibleOverlayVisible = visible;
    },
    getInvisibleOverlayVisible: () => invisibleOverlayVisible,
    setInvisibleOverlayVisible: (visible) => {
      invisibleOverlayVisible = visible;
    },
    getOverlayWindows: () => {
      const windows: BrowserWindow[] = [];
      if (mainWindow && !mainWindow.isDestroyed()) {
        windows.push(mainWindow);
      }
      if (invisibleWindow && !invisibleWindow.isDestroyed()) {
        windows.push(invisibleWindow);
      }
      return windows;
    },
  };
}
