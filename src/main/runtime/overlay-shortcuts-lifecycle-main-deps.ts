import type {
  createRefreshOverlayShortcutsHandler,
  createRegisterOverlayShortcutsHandler,
  createSyncOverlayShortcutsHandler,
  createUnregisterOverlayShortcutsHandler,
} from './overlay-shortcuts-lifecycle';

type RegisterOverlayShortcutsMainDeps = Parameters<typeof createRegisterOverlayShortcutsHandler>[0];
type UnregisterOverlayShortcutsMainDeps = Parameters<
  typeof createUnregisterOverlayShortcutsHandler
>[0];
type SyncOverlayShortcutsMainDeps = Parameters<typeof createSyncOverlayShortcutsHandler>[0];
type RefreshOverlayShortcutsMainDeps = Parameters<typeof createRefreshOverlayShortcutsHandler>[0];

export function createBuildRegisterOverlayShortcutsMainDepsHandler(
  deps: RegisterOverlayShortcutsMainDeps,
) {
  return (): RegisterOverlayShortcutsMainDeps => ({
    overlayShortcutsRuntime: deps.overlayShortcutsRuntime,
  });
}

export function createBuildUnregisterOverlayShortcutsMainDepsHandler(
  deps: UnregisterOverlayShortcutsMainDeps,
) {
  return (): UnregisterOverlayShortcutsMainDeps => ({
    overlayShortcutsRuntime: deps.overlayShortcutsRuntime,
  });
}

export function createBuildSyncOverlayShortcutsMainDepsHandler(deps: SyncOverlayShortcutsMainDeps) {
  return (): SyncOverlayShortcutsMainDeps => ({
    overlayShortcutsRuntime: deps.overlayShortcutsRuntime,
  });
}

export function createBuildRefreshOverlayShortcutsMainDepsHandler(
  deps: RefreshOverlayShortcutsMainDeps,
) {
  return (): RefreshOverlayShortcutsMainDeps => ({
    overlayShortcutsRuntime: deps.overlayShortcutsRuntime,
  });
}
