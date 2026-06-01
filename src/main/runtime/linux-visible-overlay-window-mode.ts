export type LinuxVisibleOverlayWindowMode = 'managed' | 'fullscreen-override';

type LinuxVisibleOverlayGeometry = {
  x: number;
  y: number;
  width: number;
  height: number;
};

export type LinuxVisibleOverlayWindowModeAction = {
  nextMode: LinuxVisibleOverlayWindowMode;
  shouldCreateWindow: boolean;
  shouldDestroyCurrentWindow: boolean;
  shouldRefreshVisibleOverlay: boolean;
  createWindowTiming: 'none' | 'now' | 'after-current-destroyed';
};

export function resolveLinuxVisibleOverlayWindowModeAction(options: {
  currentMode: LinuxVisibleOverlayWindowMode;
  fullscreen: boolean;
  hasLiveWindow: boolean;
  visibleOverlayVisible: boolean;
}): LinuxVisibleOverlayWindowModeAction {
  const nextMode: LinuxVisibleOverlayWindowMode = options.fullscreen
    ? 'fullscreen-override'
    : 'managed';
  const modeChanged = options.currentMode !== nextMode;

  if (!options.visibleOverlayVisible) {
    return {
      nextMode,
      shouldCreateWindow: false,
      shouldDestroyCurrentWindow: options.hasLiveWindow && modeChanged,
      shouldRefreshVisibleOverlay: false,
      createWindowTiming: 'none',
    };
  }

  if (options.hasLiveWindow && !modeChanged) {
    return {
      nextMode,
      shouldCreateWindow: false,
      shouldDestroyCurrentWindow: false,
      shouldRefreshVisibleOverlay: false,
      createWindowTiming: 'none',
    };
  }

  return {
    nextMode,
    shouldCreateWindow: true,
    shouldDestroyCurrentWindow: options.hasLiveWindow,
    shouldRefreshVisibleOverlay: true,
    createWindowTiming: options.hasLiveWindow ? 'after-current-destroyed' : 'now',
  };
}

function geometryCoversDisplayBounds(
  geometry: LinuxVisibleOverlayGeometry,
  displayBounds: LinuxVisibleOverlayGeometry,
  tolerancePx: number,
): boolean {
  const geometryRight = geometry.x + geometry.width;
  const geometryBottom = geometry.y + geometry.height;
  const displayRight = displayBounds.x + displayBounds.width;
  const displayBottom = displayBounds.y + displayBounds.height;

  return (
    geometry.x <= displayBounds.x + tolerancePx &&
    geometry.y <= displayBounds.y + tolerancePx &&
    geometryRight >= displayRight - tolerancePx &&
    geometryBottom >= displayBottom - tolerancePx
  );
}

export function shouldExitFullscreenOverrideForTrackedGeometry(options: {
  currentMode: LinuxVisibleOverlayWindowMode;
  trackedFullscreen: boolean;
  geometry: LinuxVisibleOverlayGeometry;
  displayBounds: LinuxVisibleOverlayGeometry;
  tolerancePx?: number;
}): boolean {
  if (options.currentMode !== 'fullscreen-override') return false;
  if (!options.trackedFullscreen) return false;
  return !geometryCoversDisplayBounds(
    options.geometry,
    options.displayBounds,
    options.tolerancePx ?? 2,
  );
}
