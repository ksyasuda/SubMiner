import { BrowserWindow } from 'electron';
import { RuntimeOptionState, WindowGeometry } from '../../types';
import { updateOverlayWindowBounds } from './overlay-window';

type OverlayLayer = 'visible' | 'invisible';

export interface OverlayManager {
  getMainWindow: () => BrowserWindow | null;
  setMainWindow: (window: BrowserWindow | null) => void;
  getInvisibleWindow: () => BrowserWindow | null;
  setInvisibleWindow: (window: BrowserWindow | null) => void;
  getSecondaryWindow: () => BrowserWindow | null;
  setSecondaryWindow: (window: BrowserWindow | null) => void;
  getOverlayWindow: (layer: OverlayLayer) => BrowserWindow | null;
  setOverlayWindowBounds: (layer: OverlayLayer, geometry: WindowGeometry) => void;
  setSecondaryWindowBounds: (geometry: WindowGeometry) => void;
  getVisibleOverlayVisible: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  getInvisibleOverlayVisible: () => boolean;
  setInvisibleOverlayVisible: (visible: boolean) => void;
  getOverlayWindows: () => BrowserWindow[];
  broadcastToOverlayWindows: (channel: string, ...args: unknown[]) => void;
}

export function createOverlayManager(): OverlayManager {
  let mainWindow: BrowserWindow | null = null;
  let invisibleWindow: BrowserWindow | null = null;
  let secondaryWindow: BrowserWindow | null = null;
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
    getSecondaryWindow: () => secondaryWindow,
    setSecondaryWindow: (window) => {
      secondaryWindow = window;
    },
    getOverlayWindow: (layer) => (layer === 'visible' ? mainWindow : invisibleWindow),
    setOverlayWindowBounds: (layer, geometry) => {
      updateOverlayWindowBounds(geometry, layer === 'visible' ? mainWindow : invisibleWindow);
    },
    setSecondaryWindowBounds: (geometry) => {
      updateOverlayWindowBounds(geometry, secondaryWindow);
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
      if (secondaryWindow && !secondaryWindow.isDestroyed()) {
        windows.push(secondaryWindow);
      }
      return windows;
    },
    broadcastToOverlayWindows: (channel, ...args) => {
      const windows: BrowserWindow[] = [];
      if (mainWindow && !mainWindow.isDestroyed()) {
        windows.push(mainWindow);
      }
      if (invisibleWindow && !invisibleWindow.isDestroyed()) {
        windows.push(invisibleWindow);
      }
      if (secondaryWindow && !secondaryWindow.isDestroyed()) {
        windows.push(secondaryWindow);
      }
      for (const window of windows) {
        window.webContents.send(channel, ...args);
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
  broadcastToOverlayWindows: (channel: string, ...args: unknown[]) => void,
): boolean {
  if (currentEnabled === nextEnabled) return false;
  setState(nextEnabled);
  broadcastToOverlayWindows('overlay-debug-visualization:set', nextEnabled);
  return true;
}
