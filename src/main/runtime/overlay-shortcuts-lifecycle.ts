type OverlayShortcutsRuntimeLike = {
  registerOverlayShortcuts: () => void;
  unregisterOverlayShortcuts: () => void;
  syncOverlayShortcuts: () => void;
  refreshOverlayShortcuts: () => void;
};

export function createRegisterOverlayShortcutsHandler(deps: {
  overlayShortcutsRuntime: OverlayShortcutsRuntimeLike;
}) {
  return (): void => {
    deps.overlayShortcutsRuntime.registerOverlayShortcuts();
  };
}

export function createUnregisterOverlayShortcutsHandler(deps: {
  overlayShortcutsRuntime: OverlayShortcutsRuntimeLike;
}) {
  return (): void => {
    deps.overlayShortcutsRuntime.unregisterOverlayShortcuts();
  };
}

export function createSyncOverlayShortcutsHandler(deps: {
  overlayShortcutsRuntime: OverlayShortcutsRuntimeLike;
}) {
  return (): void => {
    deps.overlayShortcutsRuntime.syncOverlayShortcuts();
  };
}

export function createRefreshOverlayShortcutsHandler(deps: {
  overlayShortcutsRuntime: OverlayShortcutsRuntimeLike;
}) {
  return (): void => {
    deps.overlayShortcutsRuntime.refreshOverlayShortcuts();
  };
}
