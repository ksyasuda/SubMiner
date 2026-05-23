type StatsOverlayVisibilityWindow = {
  isDestroyed: () => boolean;
  isVisible: () => boolean;
  setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => void;
};

function makeOverlayMousePassive(window: StatsOverlayVisibilityWindow | null): void {
  if (!window || window.isDestroyed() || !window.isVisible()) {
    return;
  }
  window.setIgnoreMouseEvents(true, { forward: true });
}

export function createStatsOverlayVisibilityChangeHandler(deps: {
  setStatsOverlayVisibleState: (visible: boolean) => void;
  resetVisibleOverlayInteraction: () => void;
  getMainWindow: () => StatsOverlayVisibilityWindow | null;
  updateVisibleOverlayVisibility: () => void;
}) {
  return (visible: boolean): void => {
    deps.setStatsOverlayVisibleState(visible);
    deps.resetVisibleOverlayInteraction();

    if (visible) {
      makeOverlayMousePassive(deps.getMainWindow());
      deps.updateVisibleOverlayVisibility();
      return;
    }

    deps.updateVisibleOverlayVisibility();
    makeOverlayMousePassive(deps.getMainWindow());
  };
}
