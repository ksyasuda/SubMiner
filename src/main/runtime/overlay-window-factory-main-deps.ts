export function createBuildCreateOverlayWindowMainDepsHandler<TWindow>(deps: {
  createOverlayWindowCore: (
    kind: 'visible' | 'modal',
    options: {
      isDev: boolean;
      ensureOverlayWindowLevel: (window: TWindow) => void;
      onRuntimeOptionsChanged: () => void;
      setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
      isOverlayVisible: (windowKind: 'visible' | 'modal') => boolean;
      tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
      forwardTabToMpv: () => void;
      onWindowClosed: (windowKind: 'visible' | 'modal') => void;
    },
  ) => TWindow;
  isDev: boolean;
  ensureOverlayWindowLevel: (window: TWindow) => void;
  onRuntimeOptionsChanged: () => void;
  setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
  isOverlayVisible: (windowKind: 'visible' | 'modal') => boolean;
  tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
  forwardTabToMpv: () => void;
  onWindowClosed: (windowKind: 'visible' | 'modal') => void;
}) {
  return () => ({
    createOverlayWindowCore: deps.createOverlayWindowCore,
    isDev: deps.isDev,
    ensureOverlayWindowLevel: deps.ensureOverlayWindowLevel,
    onRuntimeOptionsChanged: deps.onRuntimeOptionsChanged,
    setOverlayDebugVisualizationEnabled: deps.setOverlayDebugVisualizationEnabled,
    isOverlayVisible: deps.isOverlayVisible,
    tryHandleOverlayShortcutLocalFallback: deps.tryHandleOverlayShortcutLocalFallback,
    forwardTabToMpv: deps.forwardTabToMpv,
    onWindowClosed: deps.onWindowClosed,
  });
}

export function createBuildCreateMainWindowMainDepsHandler<TWindow>(deps: {
  createOverlayWindow: (kind: 'visible' | 'modal') => TWindow;
  setMainWindow: (window: TWindow | null) => void;
}) {
  return () => ({
    createOverlayWindow: deps.createOverlayWindow,
    setMainWindow: deps.setMainWindow,
  });
}

export function createBuildCreateModalWindowMainDepsHandler<TWindow>(deps: {
  createOverlayWindow: (kind: 'visible' | 'modal') => TWindow;
  setModalWindow: (window: TWindow | null) => void;
}) {
  return () => ({
    createOverlayWindow: deps.createOverlayWindow,
    setModalWindow: deps.setModalWindow,
  });
}
