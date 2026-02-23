export function createBuildCreateOverlayWindowMainDepsHandler<TWindow>(deps: {
  createOverlayWindowCore: (
    kind: 'visible' | 'invisible' | 'secondary',
    options: {
      isDev: boolean;
      overlayDebugVisualizationEnabled: boolean;
      ensureOverlayWindowLevel: (window: TWindow) => void;
      onRuntimeOptionsChanged: () => void;
      setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
      isOverlayVisible: (windowKind: 'visible' | 'invisible' | 'secondary') => boolean;
      tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
      onWindowClosed: (windowKind: 'visible' | 'invisible' | 'secondary') => void;
    },
  ) => TWindow;
  isDev: boolean;
  getOverlayDebugVisualizationEnabled: () => boolean;
  ensureOverlayWindowLevel: (window: TWindow) => void;
  onRuntimeOptionsChanged: () => void;
  setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
  isOverlayVisible: (windowKind: 'visible' | 'invisible' | 'secondary') => boolean;
  tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
  onWindowClosed: (windowKind: 'visible' | 'invisible' | 'secondary') => void;
}) {
  return () => ({
    createOverlayWindowCore: deps.createOverlayWindowCore,
    isDev: deps.isDev,
    getOverlayDebugVisualizationEnabled: deps.getOverlayDebugVisualizationEnabled,
    ensureOverlayWindowLevel: deps.ensureOverlayWindowLevel,
    onRuntimeOptionsChanged: deps.onRuntimeOptionsChanged,
    setOverlayDebugVisualizationEnabled: deps.setOverlayDebugVisualizationEnabled,
    isOverlayVisible: deps.isOverlayVisible,
    tryHandleOverlayShortcutLocalFallback: deps.tryHandleOverlayShortcutLocalFallback,
    onWindowClosed: deps.onWindowClosed,
  });
}

export function createBuildCreateMainWindowMainDepsHandler<TWindow>(deps: {
  createOverlayWindow: (kind: 'visible' | 'invisible' | 'secondary') => TWindow;
  setMainWindow: (window: TWindow | null) => void;
}) {
  return () => ({
    createOverlayWindow: deps.createOverlayWindow,
    setMainWindow: deps.setMainWindow,
  });
}

export function createBuildCreateInvisibleWindowMainDepsHandler<TWindow>(deps: {
  createOverlayWindow: (kind: 'visible' | 'invisible' | 'secondary') => TWindow;
  setInvisibleWindow: (window: TWindow | null) => void;
}) {
  return () => ({
    createOverlayWindow: deps.createOverlayWindow,
    setInvisibleWindow: deps.setInvisibleWindow,
  });
}

export function createBuildCreateSecondaryWindowMainDepsHandler<TWindow>(deps: {
  createOverlayWindow: (kind: 'visible' | 'invisible' | 'secondary') => TWindow;
  setSecondaryWindow: (window: TWindow | null) => void;
}) {
  return () => ({
    createOverlayWindow: deps.createOverlayWindow,
    setSecondaryWindow: deps.setSecondaryWindow,
  });
}
