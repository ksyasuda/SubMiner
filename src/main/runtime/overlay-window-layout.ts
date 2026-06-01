import type { WindowGeometry } from '../../types';

type OverlayBoundsWindow = {
  isDestroyed: () => boolean;
  getBounds: () => WindowGeometry;
  getContentBounds?: () => WindowGeometry;
};

function sameGeometry(a: WindowGeometry | null | undefined, b: WindowGeometry): boolean {
  return a?.x === b.x && a.y === b.y && a.width === b.width && a.height === b.height;
}

function getWindowAlignmentBounds(window: OverlayBoundsWindow): WindowGeometry | null {
  try {
    return window.getContentBounds?.() ?? window.getBounds();
  } catch {
    try {
      return window.getBounds();
    } catch {
      return null;
    }
  }
}

export function hasLiveOverlayWindowBoundsMismatch(
  windows: Array<OverlayBoundsWindow | null | undefined>,
  geometry: WindowGeometry,
): boolean {
  return windows.some((window) => {
    if (!window || window.isDestroyed()) {
      return false;
    }
    return !sameGeometry(getWindowAlignmentBounds(window), geometry);
  });
}

export function createUpdateVisibleOverlayBoundsHandler(deps: {
  getCurrentOverlayWindowBounds?: () => WindowGeometry | null;
  shouldRefreshUnchangedGeometry?: (geometry: WindowGeometry) => boolean;
  setOverlayWindowBounds: (geometry: WindowGeometry) => void;
  afterSetOverlayWindowBounds?: (geometry: WindowGeometry) => void;
}) {
  return (geometry: WindowGeometry): void => {
    if (
      sameGeometry(deps.getCurrentOverlayWindowBounds?.(), geometry) &&
      deps.shouldRefreshUnchangedGeometry?.(geometry) !== true
    ) {
      return;
    }
    deps.setOverlayWindowBounds(geometry);
    deps.afterSetOverlayWindowBounds?.(geometry);
  };
}

export function createEnsureOverlayWindowLevelHandler(deps: {
  shouldSuppressOverlayWindowLevel?: (window: unknown) => boolean;
  ensureOverlayWindowLevelCore: (window: unknown) => void;
  afterEnsureOverlayWindowLevel?: (window: unknown) => void;
}) {
  return (window: unknown): void => {
    if (deps.shouldSuppressOverlayWindowLevel?.(window) === true) {
      return;
    }
    deps.ensureOverlayWindowLevelCore(window);
    deps.afterEnsureOverlayWindowLevel?.(window);
  };
}

export function createEnforceOverlayLayerOrderHandler(deps: {
  enforceOverlayLayerOrderCore: (params: {
    visibleOverlayVisible: boolean;
    mainWindow: unknown;
    ensureOverlayWindowLevel: (window: unknown) => void;
  }) => void;
  getVisibleOverlayVisible: () => boolean;
  getMainWindow: () => unknown;
  ensureOverlayWindowLevel: (window: unknown) => void;
}) {
  return (): void => {
    deps.enforceOverlayLayerOrderCore({
      visibleOverlayVisible: deps.getVisibleOverlayVisible(),
      mainWindow: deps.getMainWindow(),
      ensureOverlayWindowLevel: deps.ensureOverlayWindowLevel,
    });
  };
}
