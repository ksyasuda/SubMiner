import { BrowserWindow } from "electron";
import { RuntimeOptionState } from "../../types";

export function getOverlayWindowsRuntimeService(options: {
  mainWindow: BrowserWindow | null;
  invisibleWindow: BrowserWindow | null;
}): BrowserWindow[] {
  const windows: BrowserWindow[] = [];
  if (options.mainWindow && !options.mainWindow.isDestroyed()) {
    windows.push(options.mainWindow);
  }
  if (options.invisibleWindow && !options.invisibleWindow.isDestroyed()) {
    windows.push(options.invisibleWindow);
  }
  return windows;
}

export function broadcastToOverlayWindowsRuntimeService(
  windows: BrowserWindow[],
  channel: string,
  ...args: unknown[]
): void {
  for (const window of windows) {
    window.webContents.send(channel, ...args);
  }
}

export function broadcastRuntimeOptionsChangedRuntimeService(
  getRuntimeOptionsState: () => RuntimeOptionState[],
  broadcastToOverlayWindows: (channel: string, ...args: unknown[]) => void,
): void {
  broadcastToOverlayWindows("runtime-options:changed", getRuntimeOptionsState());
}

export function setOverlayDebugVisualizationEnabledRuntimeService(
  currentEnabled: boolean,
  nextEnabled: boolean,
  setState: (enabled: boolean) => void,
  broadcastToOverlayWindows: (channel: string, ...args: unknown[]) => void,
): boolean {
  if (currentEnabled === nextEnabled) return false;
  setState(nextEnabled);
  broadcastToOverlayWindows("overlay-debug-visualization:set", nextEnabled);
  return true;
}
