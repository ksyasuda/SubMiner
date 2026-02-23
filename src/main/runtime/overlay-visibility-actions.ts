export function createSetVisibleOverlayVisibleHandler(deps: {
  setVisibleOverlayVisibleCore: (options: {
    visible: boolean;
    setVisibleOverlayVisibleState: (visible: boolean) => void;
    updateVisibleOverlayVisibility: () => void;
    updateInvisibleOverlayVisibility: () => void;
    syncInvisibleOverlayMousePassthrough: () => void;
    shouldBindVisibleOverlayToMpvSubVisibility: () => boolean;
    isMpvConnected: () => boolean;
    setMpvSubVisibility: (visible: boolean) => void;
  }) => void;
  setVisibleOverlayVisibleState: (visible: boolean) => void;
  updateVisibleOverlayVisibility: () => void;
  updateInvisibleOverlayVisibility: () => void;
  syncInvisibleOverlayMousePassthrough: () => void;
  shouldBindVisibleOverlayToMpvSubVisibility: () => boolean;
  isMpvConnected: () => boolean;
  setMpvSubVisibility: (visible: boolean) => void;
}) {
  return (visible: boolean): void => {
    deps.setVisibleOverlayVisibleCore({
      visible,
      setVisibleOverlayVisibleState: deps.setVisibleOverlayVisibleState,
      updateVisibleOverlayVisibility: deps.updateVisibleOverlayVisibility,
      updateInvisibleOverlayVisibility: deps.updateInvisibleOverlayVisibility,
      syncInvisibleOverlayMousePassthrough: deps.syncInvisibleOverlayMousePassthrough,
      shouldBindVisibleOverlayToMpvSubVisibility:
        deps.shouldBindVisibleOverlayToMpvSubVisibility,
      isMpvConnected: deps.isMpvConnected,
      setMpvSubVisibility: deps.setMpvSubVisibility,
    });
  };
}

export function createSetInvisibleOverlayVisibleHandler(deps: {
  setInvisibleOverlayVisibleCore: (options: {
    visible: boolean;
    setInvisibleOverlayVisibleState: (visible: boolean) => void;
    updateInvisibleOverlayVisibility: () => void;
    syncInvisibleOverlayMousePassthrough: () => void;
  }) => void;
  setInvisibleOverlayVisibleState: (visible: boolean) => void;
  updateInvisibleOverlayVisibility: () => void;
  syncInvisibleOverlayMousePassthrough: () => void;
}) {
  return (visible: boolean): void => {
    deps.setInvisibleOverlayVisibleCore({
      visible,
      setInvisibleOverlayVisibleState: deps.setInvisibleOverlayVisibleState,
      updateInvisibleOverlayVisibility: deps.updateInvisibleOverlayVisibility,
      syncInvisibleOverlayMousePassthrough: deps.syncInvisibleOverlayMousePassthrough,
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

export function createToggleInvisibleOverlayHandler(deps: {
  getInvisibleOverlayVisible: () => boolean;
  setInvisibleOverlayVisible: (visible: boolean) => void;
}) {
  return (): void => {
    deps.setInvisibleOverlayVisible(!deps.getInvisibleOverlayVisible());
  };
}
