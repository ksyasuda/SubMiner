import type { BrowserWindow, BrowserWindowConstructorOptions } from 'electron';
import type { WindowGeometry } from '../../types';

const DEFAULT_STATS_WINDOW_WIDTH = 900;
const DEFAULT_STATS_WINDOW_HEIGHT = 700;
export const STATS_WINDOW_TITLE = 'SubMiner Stats';

type StatsWindowLevelController = Pick<BrowserWindow, 'setAlwaysOnTop' | 'moveTop'> &
  Partial<Pick<BrowserWindow, 'setVisibleOnAllWorkspaces' | 'setFullScreenable'>>;

type StatsWindowBoundsController = Pick<BrowserWindow, 'getBounds' | 'getContentBounds'>;
type StatsWindowPresentationController = Pick<BrowserWindow, 'show' | 'focus'> &
  Partial<Pick<BrowserWindow, 'showInactive'>>;

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
    title: STATS_WINDOW_TITLE,
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
    backgroundColor: '#24273a',
    show: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: options.preloadPath,
      sandbox: true,
    },
  };
}

export function resolveStatsWindowOuterBoundsForContent(
  window: StatsWindowBoundsController,
  target: WindowGeometry,
): WindowGeometry {
  const outer = window.getBounds();
  const content = window.getContentBounds();
  const leftInset = content.x - outer.x;
  const topInset = content.y - outer.y;
  const rightInset = outer.x + outer.width - (content.x + content.width);
  const bottomInset = outer.y + outer.height - (content.y + content.height);
  const insets = [leftInset, topInset, rightInset, bottomInset];

  if (insets.some((inset) => !Number.isFinite(inset) || inset < 0)) {
    return target;
  }

  return {
    x: target.x - leftInset,
    y: target.y - topInset,
    width: target.width + leftInset + rightInset,
    height: target.height + topInset + bottomInset,
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

export function presentStatsWindow(
  window: StatsWindowPresentationController,
  platform: NodeJS.Platform = process.platform,
): void {
  if (platform === 'darwin') {
    if (window.showInactive) {
      window.showInactive();
    } else {
      window.show();
    }
    return;
  }

  window.show();
  window.focus();
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
