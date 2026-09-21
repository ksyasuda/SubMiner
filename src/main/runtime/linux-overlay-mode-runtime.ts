import type { BrowserWindow } from 'electron';
import {
  resolveLinuxVisibleOverlayWindowModeAction,
  type LinuxVisibleOverlayWindowMode,
} from './linux-visible-overlay-window-mode';

type OverlayWindow = Pick<BrowserWindow, 'isDestroyed' | 'hide' | 'destroy'> & {
  once: (event: 'closed', listener: () => void) => unknown;
};

export function createLinuxOverlayModeRuntime<Window extends OverlayWindow>(deps: {
  isEnabled: () => boolean;
  isVisible: () => boolean;
  getWindow: () => Window | null;
  clearWindow: () => void;
  createWindow: () => void;
  refreshWindow: () => void;
  now: () => number;
  logDebug: (message: string) => void;
}) {
  let mode: LinuxVisibleOverlayWindowMode = 'managed';
  let fullscreen = false;
  let fullscreenChangedAtMs = 0;
  let ownerBindingKey: string | null = null;
  let generation = 0;

  function createWindowForMode(token: number, nextFullscreen: boolean): void {
    if (token !== generation || !deps.isVisible()) return;
    const existing = deps.getWindow();
    if (existing && !existing.isDestroyed()) return;
    deps.createWindow();
    deps.refreshWindow();
    deps.logDebug(
      `Switched Linux visible overlay window mode to ${mode} for mpv fullscreen=${nextFullscreen}`,
    );
  }

  function sync(nextFullscreen: boolean): void {
    if (!deps.isEnabled()) return;
    if (fullscreen !== nextFullscreen) fullscreenChangedAtMs = deps.now();
    fullscreen = nextFullscreen;
    const current = deps.getWindow();
    const action = resolveLinuxVisibleOverlayWindowModeAction({
      currentMode: mode,
      fullscreen,
      hasLiveWindow: Boolean(current && !current.isDestroyed()),
      visibleOverlayVisible: deps.isVisible(),
    });
    mode = action.nextMode;
    ownerBindingKey = null;
    const token = ++generation;
    if (!action.shouldCreateWindow && !action.shouldDestroyCurrentWindow) return;

    if (action.shouldDestroyCurrentWindow && current && !current.isDestroyed()) {
      current.once('closed', () => {
        if (deps.getWindow() === current) deps.clearWindow();
        if (action.createWindowTiming === 'after-current-destroyed') {
          createWindowForMode(token, nextFullscreen);
        }
      });
      current.hide();
      current.destroy();
    }
    if (!action.shouldCreateWindow) {
      deps.logDebug(
        `Recorded Linux visible overlay window mode ${action.nextMode} for hidden mpv fullscreen=${fullscreen}`,
      );
      return;
    }
    if (action.createWindowTiming === 'now') createWindowForMode(token, nextFullscreen);
  }

  return {
    get mode() {
      return mode;
    },
    get fullscreen() {
      return fullscreen;
    },
    get fullscreenChangedAtMs() {
      return fullscreenChangedAtMs;
    },
    get ownerBindingKey() {
      return ownerBindingKey;
    },
    set ownerBindingKey(key: string | null) {
      ownerBindingKey = key;
    },
    sync,
    cancelPendingTransition: () => {
      generation += 1;
    },
  };
}
