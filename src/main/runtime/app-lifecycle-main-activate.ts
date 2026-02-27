export function createBuildShouldRestoreWindowsOnActivateMainDepsHandler(deps: {
  isOverlayRuntimeInitialized: () => boolean;
  getAllWindowCount: () => number;
}) {
  return () => ({
    isOverlayRuntimeInitialized: () => deps.isOverlayRuntimeInitialized(),
    getAllWindowCount: () => deps.getAllWindowCount(),
  });
}

export function createBuildRestoreWindowsOnActivateMainDepsHandler(deps: {
  createMainWindow: () => void;
  updateVisibleOverlayVisibility: () => void;
  syncOverlayMpvSubtitleSuppression: () => void;
}) {
  return () => ({
    createMainWindow: () => deps.createMainWindow(),
    updateVisibleOverlayVisibility: () => deps.updateVisibleOverlayVisibility(),
    syncOverlayMpvSubtitleSuppression: () => deps.syncOverlayMpvSubtitleSuppression(),
  });
}
