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
  createInvisibleWindow: () => void;
  updateVisibleOverlayVisibility: () => void;
  updateInvisibleOverlayVisibility: () => void;
}) {
  return () => ({
    createMainWindow: () => deps.createMainWindow(),
    createInvisibleWindow: () => deps.createInvisibleWindow(),
    updateVisibleOverlayVisibility: () => deps.updateVisibleOverlayVisibility(),
    updateInvisibleOverlayVisibility: () => deps.updateInvisibleOverlayVisibility(),
  });
}
