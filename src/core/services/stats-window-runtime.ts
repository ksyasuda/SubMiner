import type { BrowserWindow, BrowserWindowConstructorOptions } from 'electron';
import type { WindowGeometry } from '../../types';

const DEFAULT_STATS_WINDOW_WIDTH = 900;
const DEFAULT_STATS_WINDOW_HEIGHT = 700;

type StatsWindowLevelController = Pick<BrowserWindow, 'setAlwaysOnTop' | 'moveTop'> &
  Partial<Pick<BrowserWindow, 'setVisibleOnAllWorkspaces' | 'setFullScreenable'>>;

function isBareToggleKeyInput(input: Electron.Input, toggleKey: string): boolean {
  return (
    input.type === 'keyDown' &&
    input.code === toggleKey &&
    !input.control &&
    !input.alt &&
    !input.meta &&
    !input.shift &&
    !input.isAutoRepeat
  );
}

export function shouldHideStatsWindowForInput(input: Electron.Input, toggleKey: string): boolean {
  return (
    (input.type === 'keyDown' && input.key === 'Escape') || isBareToggleKeyInput(input, toggleKey)
  );
}

export function buildStatsWindowOptions(options: {
  preloadPath: string;
  bounds?: WindowGeometry | null;
}): BrowserWindowConstructorOptions {
  return {
    x: options.bounds?.x,
    y: options.bounds?.y,
    width: options.bounds?.width ?? DEFAULT_STATS_WINDOW_WIDTH,
    height: options.bounds?.height ?? DEFAULT_STATS_WINDOW_HEIGHT,
    frame: false,
    transparent: true,
    alwaysOnTop: true,
    resizable: false,
    skipTaskbar: true,
    hasShadow: false,
    focusable: true,
    acceptFirstMouse: true,
    fullscreenable: false,
    backgroundColor: '#1e1e2e',
    show: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: options.preloadPath,
      sandbox: true,
    },
  };
}

export function promoteStatsWindowLevel(
  window: StatsWindowLevelController,
  platform: NodeJS.Platform = process.platform,
): void {
  if (platform === 'darwin') {
    window.setAlwaysOnTop(true, 'screen-saver', 2);
    window.setVisibleOnAllWorkspaces?.(true, { visibleOnFullScreen: true });
    window.setFullScreenable?.(false);
    window.moveTop();
    return;
  }

  if (platform === 'win32') {
    window.setAlwaysOnTop(true, 'screen-saver', 2);
    window.moveTop();
    return;
  }

  window.setAlwaysOnTop(true);
  window.moveTop();
}

export function buildStatsWindowLoadFileOptions(apiBaseUrl?: string): {
  query: Record<string, string>;
} {
  return {
    query: {
      overlay: '1',
      ...(apiBaseUrl ? { apiBase: apiBaseUrl } : {}),
    },
  };
}
