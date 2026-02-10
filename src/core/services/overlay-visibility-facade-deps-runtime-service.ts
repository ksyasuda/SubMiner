import {
  OverlayVisibilityFacadeDeps,
} from "./overlay-visibility-facade-service";

export interface OverlayVisibilityFacadeDepsRuntimeOptions {
  getVisibleOverlayVisible: () => boolean;
  getInvisibleOverlayVisible: () => boolean;
  setVisibleOverlayVisibleState: (nextVisible: boolean) => void;
  setInvisibleOverlayVisibleState: (nextVisible: boolean) => void;
  updateVisibleOverlayVisibility: () => void;
  updateInvisibleOverlayVisibility: () => void;
  syncInvisibleOverlayMousePassthrough: () => void;
  shouldBindVisibleOverlayToMpvSubVisibility: () => boolean;
  isMpvConnected: () => boolean;
  setMpvSubVisibility: (mpvSubVisible: boolean) => void;
}

export function createOverlayVisibilityFacadeDepsRuntimeService(
  options: OverlayVisibilityFacadeDepsRuntimeOptions,
): OverlayVisibilityFacadeDeps {
  return {
    getVisibleOverlayVisible: options.getVisibleOverlayVisible,
    getInvisibleOverlayVisible: options.getInvisibleOverlayVisible,
    setVisibleOverlayVisibleState: options.setVisibleOverlayVisibleState,
    setInvisibleOverlayVisibleState: options.setInvisibleOverlayVisibleState,
    updateVisibleOverlayVisibility: options.updateVisibleOverlayVisibility,
    updateInvisibleOverlayVisibility: options.updateInvisibleOverlayVisibility,
    syncInvisibleOverlayMousePassthrough:
      options.syncInvisibleOverlayMousePassthrough,
    shouldBindVisibleOverlayToMpvSubVisibility:
      options.shouldBindVisibleOverlayToMpvSubVisibility,
    isMpvConnected: options.isMpvConnected,
    setMpvSubVisibility: options.setMpvSubVisibility,
  };
}
