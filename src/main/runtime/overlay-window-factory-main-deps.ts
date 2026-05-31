import type { Session } from 'electron';

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
      linuxX11FullscreenOverlay?: boolean;
      onVisibleWindowBlurred?: () => void;
      onVisibleWindowFocused?: () => void;
      onWindowContentReady?: () => void;
      onWindowClosed: (windowKind: 'visible' | 'modal', window: TWindow) => void;
      yomitanSession?: Session | null;
    },
  ) => TWindow;
  isDev: boolean;
  ensureOverlayWindowLevel: (window: TWindow) => void;
  onRuntimeOptionsChanged: () => void;
  setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
  isOverlayVisible: (windowKind: 'visible' | 'modal') => boolean;
  tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
  forwardTabToMpv: () => void;
  getLinuxX11FullscreenOverlay?: () => boolean;
  onVisibleWindowBlurred?: () => void;
  onVisibleWindowFocused?: () => void;
  onWindowContentReady?: () => void;
  onWindowClosed: (windowKind: 'visible' | 'modal', window: TWindow) => void;
  getYomitanSession?: () => Session | null;
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
    getLinuxX11FullscreenOverlay: deps.getLinuxX11FullscreenOverlay,
    onVisibleWindowBlurred: deps.onVisibleWindowBlurred,
    onVisibleWindowFocused: deps.onVisibleWindowFocused,
    onWindowContentReady: deps.onWindowContentReady,
    onWindowClosed: deps.onWindowClosed,
    getYomitanSession: () => deps.getYomitanSession?.() ?? null,
  });
}

export function createBuildCreateMainWindowMainDepsHandler<TWindow>(deps: {
  getMainWindow: () => TWindow | null;
  isWindowDestroyed: (window: TWindow) => boolean;
  createOverlayWindow: (kind: 'visible' | 'modal') => TWindow;
  setMainWindow: (window: TWindow | null) => void;
}) {
  return () => ({
    getMainWindow: () => deps.getMainWindow(),
    isWindowDestroyed: (window: TWindow) => deps.isWindowDestroyed(window),
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
