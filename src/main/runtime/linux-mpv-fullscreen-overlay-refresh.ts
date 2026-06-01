type LinuxMpvFullscreenOverlayWindow = {
  hide: () => void;
  isDestroyed: () => boolean;
  isVisible: () => boolean;
  setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => void;
  showInactive: () => void;
};

export type LinuxMpvFullscreenOverlayRefreshDeps = {
  overlayManager: {
    getMainWindow: () => LinuxMpvFullscreenOverlayWindow | null;
    getVisibleOverlayVisible: () => boolean;
  };
  overlayVisibilityRuntime: {
    updateVisibleOverlayVisibility: () => void;
  };
  syncVisibleOverlayMpvFullscreenMode?: (fullscreen: boolean) => void;
  getOverlayInteractionActive?: () => boolean;
  ensureOverlayWindowLevel: (window: LinuxMpvFullscreenOverlayWindow) => void;
};
export type CancelLinuxMpvFullscreenOverlayRefreshBurst = () => void;

const LINUX_MPV_FULLSCREEN_OVERLAY_REFRESH_DELAYS_MS = [0, 50, 150, 300, 600] as const;
let linuxMpvFullscreenOverlayRefreshTimeouts: Array<ReturnType<typeof setTimeout>> = [];

function clearLinuxMpvFullscreenOverlayRefreshTimeouts(): void {
  for (const timeout of linuxMpvFullscreenOverlayRefreshTimeouts) {
    clearTimeout(timeout);
  }
  linuxMpvFullscreenOverlayRefreshTimeouts = [];
}

function refreshLinuxVisibleOverlayAfterMpvFullscreenChange(
  fullscreen: boolean,
  deps: LinuxMpvFullscreenOverlayRefreshDeps,
): void {
  if (process.platform !== 'linux') {
    return;
  }

  deps.syncVisibleOverlayMpvFullscreenMode?.(fullscreen);
  if (!deps.overlayManager.getVisibleOverlayVisible()) {
    return;
  }
  deps.overlayVisibilityRuntime.updateVisibleOverlayVisibility();
  if (!fullscreen) {
    return;
  }

  const mainWindow = deps.overlayManager.getMainWindow();
  if (!mainWindow || mainWindow.isDestroyed() || !mainWindow.isVisible()) {
    return;
  }

  mainWindow.hide();
  mainWindow.showInactive();
  if (deps.getOverlayInteractionActive?.() === true) {
    mainWindow.setIgnoreMouseEvents(false);
  } else {
    mainWindow.setIgnoreMouseEvents(true, { forward: true });
  }
  deps.ensureOverlayWindowLevel(mainWindow);
}

export function scheduleLinuxVisibleOverlayFullscreenRefreshBurst(
  isFullscreen: boolean,
  deps: LinuxMpvFullscreenOverlayRefreshDeps,
): CancelLinuxMpvFullscreenOverlayRefreshBurst {
  if (process.platform !== 'linux') {
    return () => {};
  }

  clearLinuxMpvFullscreenOverlayRefreshTimeouts();
  for (const delayMs of LINUX_MPV_FULLSCREEN_OVERLAY_REFRESH_DELAYS_MS) {
    const refreshTimeout = setTimeout(() => {
      linuxMpvFullscreenOverlayRefreshTimeouts = linuxMpvFullscreenOverlayRefreshTimeouts.filter(
        (timeout) => timeout !== refreshTimeout,
      );
      refreshLinuxVisibleOverlayAfterMpvFullscreenChange(isFullscreen, deps);
    }, delayMs);
    refreshTimeout.unref?.();
    linuxMpvFullscreenOverlayRefreshTimeouts.push(refreshTimeout);
  }
  return clearLinuxMpvFullscreenOverlayRefreshTimeouts;
}

export function updateLinuxMpvFullscreenOverlayRefreshBurst(
  isFullscreen: boolean,
  deps: LinuxMpvFullscreenOverlayRefreshDeps,
  cancelCurrentBurst: CancelLinuxMpvFullscreenOverlayRefreshBurst | null,
): CancelLinuxMpvFullscreenOverlayRefreshBurst | null {
  cancelCurrentBurst?.();

  return scheduleLinuxVisibleOverlayFullscreenRefreshBurst(isFullscreen, deps);
}

export { clearLinuxMpvFullscreenOverlayRefreshTimeouts };
