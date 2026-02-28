export function createSetVisibleOverlayVisibleHandler(deps: {
  setVisibleOverlayVisibleCore: (options: {
    visible: boolean;
    setVisibleOverlayVisibleState: (visible: boolean) => void;
    updateVisibleOverlayVisibility: () => void;
  }) => void;
  setVisibleOverlayVisibleState: (visible: boolean) => void;
  updateVisibleOverlayVisibility: () => void;
  onVisibleOverlayEnabled?: () => void;
}) {
  return (visible: boolean): void => {
    if (visible) {
      deps.onVisibleOverlayEnabled?.();
    }
    deps.setVisibleOverlayVisibleCore({
      visible,
      setVisibleOverlayVisibleState: deps.setVisibleOverlayVisibleState,
      updateVisibleOverlayVisibility: deps.updateVisibleOverlayVisibility,
    });
  };
}

export function createToggleVisibleOverlayHandler(deps: {
  getVisibleOverlayVisible: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
}) {
  return (): void => {
    deps.setVisibleOverlayVisible(!deps.getVisibleOverlayVisible());
  };
}
