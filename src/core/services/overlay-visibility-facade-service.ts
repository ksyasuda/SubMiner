import {
  setInvisibleOverlayVisibleService,
  setVisibleOverlayVisibleService,
} from "./overlay-visibility-runtime-service";

export interface OverlayVisibilityFacadeDeps {
  getVisibleOverlayVisible: () => boolean;
  getInvisibleOverlayVisible: () => boolean;
  setVisibleOverlayVisibleState: (visible: boolean) => void;
  setInvisibleOverlayVisibleState: (visible: boolean) => void;
  updateVisibleOverlayVisibility: () => void;
  updateInvisibleOverlayVisibility: () => void;
  syncInvisibleOverlayMousePassthrough: () => void;
  shouldBindVisibleOverlayToMpvSubVisibility: () => boolean;
  isMpvConnected: () => boolean;
  setMpvSubVisibility: (visible: boolean) => void;
}

export function setVisibleOverlayVisibleRuntimeFacadeService(
  visible: boolean,
  deps: OverlayVisibilityFacadeDeps,
): void {
  setVisibleOverlayVisibleService({
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
}

export function setInvisibleOverlayVisibleRuntimeFacadeService(
  visible: boolean,
  deps: OverlayVisibilityFacadeDeps,
): void {
  setInvisibleOverlayVisibleService({
    visible,
    setInvisibleOverlayVisibleState: deps.setInvisibleOverlayVisibleState,
    updateInvisibleOverlayVisibility: deps.updateInvisibleOverlayVisibility,
    syncInvisibleOverlayMousePassthrough: deps.syncInvisibleOverlayMousePassthrough,
  });
}

export function toggleVisibleOverlayRuntimeFacadeService(
  deps: OverlayVisibilityFacadeDeps,
): void {
  setVisibleOverlayVisibleRuntimeFacadeService(
    !deps.getVisibleOverlayVisible(),
    deps,
  );
}

export function toggleInvisibleOverlayRuntimeFacadeService(
  deps: OverlayVisibilityFacadeDeps,
): void {
  setInvisibleOverlayVisibleRuntimeFacadeService(
    !deps.getInvisibleOverlayVisible(),
    deps,
  );
}
