type OverlayWindowKind = 'visible' | 'invisible' | 'secondary';

export function createCreateOverlayWindowHandler<TWindow>(deps: {
  createOverlayWindowCore: (
    kind: OverlayWindowKind,
    options: {
      isDev: boolean;
      overlayDebugVisualizationEnabled: boolean;
      ensureOverlayWindowLevel: (window: TWindow) => void;
      onRuntimeOptionsChanged: () => void;
      setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
      isOverlayVisible: (windowKind: OverlayWindowKind) => boolean;
      tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
      onWindowClosed: (windowKind: OverlayWindowKind) => void;
    },
  ) => TWindow;
  isDev: boolean;
  getOverlayDebugVisualizationEnabled: () => boolean;
  ensureOverlayWindowLevel: (window: TWindow) => void;
  onRuntimeOptionsChanged: () => void;
  setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
  isOverlayVisible: (windowKind: OverlayWindowKind) => boolean;
  tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
  onWindowClosed: (windowKind: OverlayWindowKind) => void;
}) {
  return (kind: OverlayWindowKind): TWindow => {
    return deps.createOverlayWindowCore(kind, {
      isDev: deps.isDev,
      overlayDebugVisualizationEnabled: deps.getOverlayDebugVisualizationEnabled(),
      ensureOverlayWindowLevel: deps.ensureOverlayWindowLevel,
      onRuntimeOptionsChanged: deps.onRuntimeOptionsChanged,
      setOverlayDebugVisualizationEnabled: deps.setOverlayDebugVisualizationEnabled,
      isOverlayVisible: deps.isOverlayVisible,
      tryHandleOverlayShortcutLocalFallback: deps.tryHandleOverlayShortcutLocalFallback,
      onWindowClosed: deps.onWindowClosed,
    });
  };
}

export function createCreateMainWindowHandler<TWindow>(deps: {
  createOverlayWindow: (kind: OverlayWindowKind) => TWindow;
  setMainWindow: (window: TWindow | null) => void;
}) {
  return (): TWindow => {
    const window = deps.createOverlayWindow('visible');
    deps.setMainWindow(window);
    return window;
  };
}

export function createCreateInvisibleWindowHandler<TWindow>(deps: {
  createOverlayWindow: (kind: OverlayWindowKind) => TWindow;
  setInvisibleWindow: (window: TWindow | null) => void;
}) {
  return (): TWindow => {
    const window = deps.createOverlayWindow('invisible');
    deps.setInvisibleWindow(window);
    return window;
  };
}

export function createCreateSecondaryWindowHandler<TWindow>(deps: {
  createOverlayWindow: (kind: OverlayWindowKind) => TWindow;
  setSecondaryWindow: (window: TWindow | null) => void;
}) {
  return (): TWindow => {
    const window = deps.createOverlayWindow('secondary');
    deps.setSecondaryWindow(window);
    return window;
  };
}
