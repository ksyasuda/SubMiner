export function createBuildCreateOverlayWindowMainDepsHandler<TWindow>(deps: {
  createOverlayWindowCore: (
    kind: 'visible' | 'invisible',
    options: {
      isDev: boolean;
      overlayDebugVisualizationEnabled: boolean;
      ensureOverlayWindowLevel: (window: TWindow) => void;
      onRuntimeOptionsChanged: () => void;
      setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
      isOverlayVisible: (windowKind: 'visible' | 'invisible') => boolean;
      tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
      onWindowClosed: (windowKind: 'visible' | 'invisible') => void;
    },
  ) => TWindow;
  isDev: boolean;
  getOverlayDebugVisualizationEnabled: () => boolean;
  ensureOverlayWindowLevel: (window: TWindow) => void;
  onRuntimeOptionsChanged: () => void;
  setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
  isOverlayVisible: (windowKind: 'visible' | 'invisible') => boolean;
  tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
  onWindowClosed: (windowKind: 'visible' | 'invisible') => void;
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
  createOverlayWindow: (kind: 'visible' | 'invisible') => TWindow;
  setMainWindow: (window: TWindow | null) => void;
}) {
  return () => ({
    createOverlayWindow: deps.createOverlayWindow,
    setMainWindow: deps.setMainWindow,
  });
}

export function createBuildCreateInvisibleWindowMainDepsHandler<TWindow>(deps: {
  createOverlayWindow: (kind: 'visible' | 'invisible') => TWindow;
  setInvisibleWindow: (window: TWindow | null) => void;
}) {
  return () => ({
    createOverlayWindow: deps.createOverlayWindow,
    setInvisibleWindow: deps.setInvisibleWindow,
  });
}
