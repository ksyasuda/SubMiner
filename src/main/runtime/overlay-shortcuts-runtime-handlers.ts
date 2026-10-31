import {
  createRefreshOverlayShortcutsHandler,
  createRegisterOverlayShortcutsHandler,
  createSyncOverlayShortcutsHandler,
  createUnregisterOverlayShortcutsHandler,
} from './overlay-shortcuts-lifecycle';

type RegisterOverlayShortcutsMainDeps = Parameters<typeof createRegisterOverlayShortcutsHandler>[0];

export function createOverlayShortcutsRuntimeHandlers(deps: {
  overlayShortcutsRuntimeMainDeps: RegisterOverlayShortcutsMainDeps;
}) {
  const registerOverlayShortcutsHandler = createRegisterOverlayShortcutsHandler(
    deps.overlayShortcutsRuntimeMainDeps,
  );

  const unregisterOverlayShortcutsHandler = createUnregisterOverlayShortcutsHandler(
    deps.overlayShortcutsRuntimeMainDeps,
  );

  const syncOverlayShortcutsHandler = createSyncOverlayShortcutsHandler(
    deps.overlayShortcutsRuntimeMainDeps,
  );

  const refreshOverlayShortcutsHandler = createRefreshOverlayShortcutsHandler(
    deps.overlayShortcutsRuntimeMainDeps,
  );

  return {
    registerOverlayShortcuts: () => registerOverlayShortcutsHandler(),
    unregisterOverlayShortcuts: () => unregisterOverlayShortcutsHandler(),
    syncOverlayShortcuts: () => syncOverlayShortcutsHandler(),
    refreshOverlayShortcuts: () => refreshOverlayShortcutsHandler(),
  };
}
