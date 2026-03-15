import type { BrowserWindowConstructorOptions } from 'electron';
import type { WindowGeometry } from '../../types';

const DEFAULT_STATS_WINDOW_WIDTH = 900;
const DEFAULT_STATS_WINDOW_HEIGHT = 700;

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

export function shouldHideStatsWindowForInput(
  input: Electron.Input,
  toggleKey: string,
): boolean {
  return (
    (input.type === 'keyDown' && input.key === 'Escape') ||
    isBareToggleKeyInput(input, toggleKey)
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
    transparent: false,
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
      sandbox: false,
    },
  };
}

export function buildStatsWindowLoadFileOptions(): { query: Record<string, string> } {
  return {
    query: {
      overlay: '1',
    },
  };
}
