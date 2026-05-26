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

type UpdateOverlayWindowBounds = typeof updateOverlayWindowBounds;

export interface OverlayManagerOptions {
  updateOverlayWindowBounds?: UpdateOverlayWindowBounds;
  shouldPromoteWindowOnBoundsUpdate?: (window: BrowserWindow) => boolean;
}

export function createOverlayManager(options: OverlayManagerOptions = {}): OverlayManager {
  let mainWindow: BrowserWindow | null = null;
  let modalWindow: BrowserWindow | null = null;
  let visibleOverlayVisible = false;
  const applyOverlayBounds = options.updateOverlayWindowBounds ?? updateOverlayWindowBounds;

  const updateWindowBounds = (geometry: WindowGeometry, window: BrowserWindow | null): void => {
    const promote = window ? (options.shouldPromoteWindowOnBoundsUpdate?.(window) ?? true) : true;
    applyOverlayBounds(geometry, window, { promote });
  };

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
      updateWindowBounds(geometry, mainWindow);
    },
    setModalWindowBounds: (geometry) => {
      updateWindowBounds(geometry, modalWindow);
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
