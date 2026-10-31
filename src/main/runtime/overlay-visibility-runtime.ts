import {
  createSetVisibleOverlayVisibleHandler,
  createToggleVisibleOverlayHandler,
} from './overlay-visibility-actions';
import { createSetOverlayVisibleHandler, createToggleOverlayHandler } from './overlay-main-actions';

type SetVisibleOverlayVisibleDeps = Parameters<typeof createSetVisibleOverlayVisibleHandler>[0];

export type OverlayVisibilityRuntimeDeps = {
  setVisibleOverlayVisibleDeps: SetVisibleOverlayVisibleDeps;
  getVisibleOverlayVisible: () => boolean;
};

export function createOverlayVisibilityRuntime(deps: OverlayVisibilityRuntimeDeps) {
  const setVisibleOverlayVisible = createSetVisibleOverlayVisibleHandler({
    ...deps.setVisibleOverlayVisibleDeps,
    getVisibleOverlayVisible: deps.getVisibleOverlayVisible,
  });

  const toggleVisibleOverlay = createToggleVisibleOverlayHandler({
    getVisibleOverlayVisible: deps.getVisibleOverlayVisible,
    setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
  });

  const setOverlayVisible = createSetOverlayVisibleHandler({
    setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
  });

  const toggleOverlay = createToggleOverlayHandler({
    toggleVisibleOverlay: () => toggleVisibleOverlay(),
  });

  return {
    setVisibleOverlayVisible,
    toggleVisibleOverlay,
    setOverlayVisible,
    toggleOverlay,
  };
}
