import type { BrowserWindow } from 'electron';
import { RuntimeOptionState, WindowGeometry } from '../../types';
import { updateOverlayWindowBounds } from './overlay-window';

export interface OverlayManager {
  getMainWindow: () => BrowserWindow | null;
  setMainWindow: (window: BrowserWindow | null) => void;
  getModalWindow: () => BrowserWindow | null;
  setModalWindow: (window: BrowserWindow | null) => void;
  getOverlayWindow: () => BrowserWindow | null;
  setOverlayWindowBounds: (geometry: WindowGeometry) => void;
  setModalWindowBounds: (geometry: WindowGeometry) => void;
  getVisibleOverlayVisible: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  getOverlayWindows: () => BrowserWindow[];
  broadcastToOverlayWindows: (channel: string, ...args: unknown[]) => void;
}

export function createOverlayManager(): OverlayManager {
  let mainWindow: BrowserWindow | null = null;
  let modalWindow: BrowserWindow | null = null;
  let visibleOverlayVisible = false;

  return {
    getMainWindow: () => mainWindow,
    setMainWindow: (window) => {
      mainWindow = window;
    },
    getModalWindow: () => modalWindow,
    setModalWindow: (window) => {
      modalWindow = window;
    },
    getOverlayWindow: () => mainWindow,
    setOverlayWindowBounds: (geometry) => {
      updateOverlayWindowBounds(geometry, mainWindow);
    },
    setModalWindowBounds: (geometry) => {
      updateOverlayWindowBounds(geometry, modalWindow);
    },
    getVisibleOverlayVisible: () => visibleOverlayVisible,
    setVisibleOverlayVisible: (visible) => {
      visibleOverlayVisible = visible;
    },
    getOverlayWindows: () => {
      return mainWindow && !mainWindow.isDestroyed() ? [mainWindow] : [];
    },
    broadcastToOverlayWindows: (channel, ...args) => {
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send(channel, ...args);
      }
    },
  };
}

export function broadcastRuntimeOptionsChangedRuntime(
  getRuntimeOptionsState: () => RuntimeOptionState[],
  broadcastToOverlayWindows: (channel: string, ...args: unknown[]) => void,
): void {
  broadcastToOverlayWindows('runtime-options:changed', getRuntimeOptionsState());
}

export function setOverlayDebugVisualizationEnabledRuntime(
  currentEnabled: boolean,
  nextEnabled: boolean,
  setState: (enabled: boolean) => void,
): boolean {
  if (currentEnabled === nextEnabled) return false;
  setState(nextEnabled);
  return true;
}
