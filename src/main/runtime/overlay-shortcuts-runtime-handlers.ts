import {
  createRefreshOverlayShortcutsHandler,
  createRegisterOverlayShortcutsHandler,
  createSyncOverlayShortcutsHandler,
  createUnregisterOverlayShortcutsHandler,
} from './overlay-shortcuts-lifecycle';
import {
  createBuildRefreshOverlayShortcutsMainDepsHandler,
  createBuildRegisterOverlayShortcutsMainDepsHandler,
  createBuildSyncOverlayShortcutsMainDepsHandler,
  createBuildUnregisterOverlayShortcutsMainDepsHandler,
} from './overlay-shortcuts-lifecycle-main-deps';

type RegisterOverlayShortcutsMainDeps = Parameters<
  typeof createBuildRegisterOverlayShortcutsMainDepsHandler
>[0];

export function createOverlayShortcutsRuntimeHandlers(deps: {
  overlayShortcutsRuntimeMainDeps: RegisterOverlayShortcutsMainDeps;
}) {
  const registerOverlayShortcutsMainDeps = createBuildRegisterOverlayShortcutsMainDepsHandler(
    deps.overlayShortcutsRuntimeMainDeps,
  )();
  const registerOverlayShortcutsHandler =
    createRegisterOverlayShortcutsHandler(registerOverlayShortcutsMainDeps);

  const unregisterOverlayShortcutsMainDeps =
    createBuildUnregisterOverlayShortcutsMainDepsHandler(
      deps.overlayShortcutsRuntimeMainDeps,
    )();
  const unregisterOverlayShortcutsHandler =
    createUnregisterOverlayShortcutsHandler(unregisterOverlayShortcutsMainDeps);

  const syncOverlayShortcutsMainDeps = createBuildSyncOverlayShortcutsMainDepsHandler(
    deps.overlayShortcutsRuntimeMainDeps,
  )();
  const syncOverlayShortcutsHandler = createSyncOverlayShortcutsHandler(syncOverlayShortcutsMainDeps);

  const refreshOverlayShortcutsMainDeps = createBuildRefreshOverlayShortcutsMainDepsHandler(
    deps.overlayShortcutsRuntimeMainDeps,
  )();
  const refreshOverlayShortcutsHandler =
    createRefreshOverlayShortcutsHandler(refreshOverlayShortcutsMainDeps);

  return {
    registerOverlayShortcuts: () => registerOverlayShortcutsHandler(),
    unregisterOverlayShortcuts: () => unregisterOverlayShortcutsHandler(),
    syncOverlayShortcuts: () => syncOverlayShortcutsHandler(),
    refreshOverlayShortcuts: () => refreshOverlayShortcutsHandler(),
  };
}
