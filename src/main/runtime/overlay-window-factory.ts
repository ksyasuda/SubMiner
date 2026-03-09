type OverlayWindowKind = 'visible' | 'modal';

export function createCreateOverlayWindowHandler<TWindow>(deps: {
  createOverlayWindowCore: (
    kind: OverlayWindowKind,
    options: {
      isDev: boolean;
      ensureOverlayWindowLevel: (window: TWindow) => void;
      onRuntimeOptionsChanged: () => void;
      setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
      isOverlayVisible: (windowKind: OverlayWindowKind) => boolean;
      tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
      forwardTabToMpv: () => void;
      onWindowClosed: (windowKind: OverlayWindowKind) => void;
    },
  ) => TWindow;
  isDev: boolean;
  ensureOverlayWindowLevel: (window: TWindow) => void;
  onRuntimeOptionsChanged: () => void;
  setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
  isOverlayVisible: (windowKind: OverlayWindowKind) => boolean;
  tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
  forwardTabToMpv: () => void;
  onWindowClosed: (windowKind: OverlayWindowKind) => void;
}) {
  return (kind: OverlayWindowKind): TWindow => {
    return deps.createOverlayWindowCore(kind, {
      isDev: deps.isDev,
      ensureOverlayWindowLevel: deps.ensureOverlayWindowLevel,
      onRuntimeOptionsChanged: deps.onRuntimeOptionsChanged,
      setOverlayDebugVisualizationEnabled: deps.setOverlayDebugVisualizationEnabled,
      isOverlayVisible: deps.isOverlayVisible,
      tryHandleOverlayShortcutLocalFallback: deps.tryHandleOverlayShortcutLocalFallback,
      forwardTabToMpv: deps.forwardTabToMpv,
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

export function createCreateModalWindowHandler<TWindow>(deps: {
  createOverlayWindow: (kind: OverlayWindowKind) => TWindow;
  setModalWindow: (window: TWindow | null) => void;
}) {
  return (): TWindow => {
    const window = deps.createOverlayWindow('modal');
    deps.setModalWindow(window);
    return window;
  };
}
