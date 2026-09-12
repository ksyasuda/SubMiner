import type {
  BrowserWindow,
  BrowserWindowConstructorOptions,
  MessageBoxSyncOptions,
} from 'electron';
import type { WindowGeometry } from '../../types';

const DEFAULT_STATS_WINDOW_WIDTH = 900;
const DEFAULT_STATS_WINDOW_HEIGHT = 700;
export const STATS_WINDOW_TITLE = 'SubMiner Stats';
const STATS_POST_SHOW_RECONCILE_DELAYS_MS = [50, 150, 300, 600] as const;

type StatsWindowLevelController = Pick<BrowserWindow, 'setAlwaysOnTop' | 'moveTop'> &
  Partial<Pick<BrowserWindow, 'setVisibleOnAllWorkspaces' | 'setFullScreenable'>>;
type VisibleStatsWindowLevelController = StatsWindowLevelController &
  Pick<BrowserWindow, 'isDestroyed' | 'isVisible'>;
type VisibleStatsWindowDialogLayerController = Pick<
  BrowserWindow,
  'isDestroyed' | 'isVisible' | 'setAlwaysOnTop'
>;
type StatsNativeConfirmDialogWindow = Pick<BrowserWindow, 'isDestroyed'>;
type StatsNativeConfirmDialogPresenter<WindowT> = {
  showWithParent: (window: WindowT, options: MessageBoxSyncOptions) => number;
  showWithoutParent: (options: MessageBoxSyncOptions) => number;
};

type StatsWindowBoundsController = Pick<BrowserWindow, 'getBounds' | 'getContentBounds'>;
type StatsWindowPresentationController = Pick<BrowserWindow, 'show' | 'focus'> &
  Partial<Pick<BrowserWindow, 'showInactive'>>;
type StatsWindowReconcileScheduler = {
  setTimeout: (
    callback: () => void,
    delayMs: number,
  ) => {
    unref?: () => void;
  };
};

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
  platform?: NodeJS.Platform;
}): BrowserWindowConstructorOptions {
  const platform = options.platform ?? process.platform;
  return {
    title: STATS_WINDOW_TITLE,
    x: options.bounds?.x,
    y: options.bounds?.y,
    width: options.bounds?.width ?? DEFAULT_STATS_WINDOW_WIDTH,
    height: options.bounds?.height ?? DEFAULT_STATS_WINDOW_HEIGHT,
    frame: false,
    ...(platform === 'linux' ? { roundedCorners: false } : {}),
    transparent: false,
    alwaysOnTop: true,
    resizable: false,
    skipTaskbar: true,
    hasShadow: false,
    focusable: true,
    acceptFirstMouse: true,
    fullscreenable: false,
    // Panels join fullscreen Spaces on macOS without moving the user back to the
    // desktop where SubMiner last owned a regular application window.
    ...(platform === 'darwin' ? { type: 'panel' as const } : {}),
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

export function shouldPresentStatsWindowAfterLoad(
  platform: NodeJS.Platform = process.platform,
): boolean {
  return platform === 'darwin';
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

export function promoteVisibleStatsWindowAboveOverlay(
  window: VisibleStatsWindowLevelController,
  options: {
    platform?: NodeJS.Platform;
    promoteHyprlandWindow?: () => void;
  } = {},
): boolean {
  if (window.isDestroyed() || !window.isVisible()) {
    return false;
  }

  promoteStatsWindowLevel(window, options.platform);
  options.promoteHyprlandWindow?.();
  return true;
}

export function demoteVisibleStatsWindowBelowDialogs(
  window: VisibleStatsWindowDialogLayerController,
): boolean {
  if (window.isDestroyed() || !window.isVisible()) {
    return false;
  }

  window.setAlwaysOnTop(false);
  return true;
}

export function buildStatsNativeConfirmDialogOptions(message: string): MessageBoxSyncOptions {
  return {
    type: 'warning',
    message,
    buttons: ['Delete', 'Cancel'],
    defaultId: 1,
    cancelId: 1,
    noLink: true,
  };
}

export function showStatsNativeConfirmDialog<WindowT extends StatsNativeConfirmDialogWindow>(
  window: WindowT | null,
  message: string,
  presenter: StatsNativeConfirmDialogPresenter<WindowT>,
): boolean {
  const options = buildStatsNativeConfirmDialogOptions(message);
  const response =
    window && !window.isDestroyed()
      ? presenter.showWithParent(window, options)
      : presenter.showWithoutParent(options);
  return response === 0;
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

export function scheduleStatsWindowPostShowReconciles(
  reconcile: () => void,
  scheduler: StatsWindowReconcileScheduler = globalThis,
): void {
  for (const delayMs of STATS_POST_SHOW_RECONCILE_DELAYS_MS) {
    const timeout = scheduler.setTimeout(reconcile, delayMs);
    timeout.unref?.();
  }
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
