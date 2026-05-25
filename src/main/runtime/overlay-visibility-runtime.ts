import {
  createSetVisibleOverlayVisibleHandler,
  createToggleVisibleOverlayHandler,
} from './overlay-visibility-actions';
import {
  createBuildSetVisibleOverlayVisibleMainDepsHandler,
  createBuildToggleVisibleOverlayMainDepsHandler,
} from './overlay-visibility-actions-main-deps';
import { createSetOverlayVisibleHandler, createToggleOverlayHandler } from './overlay-main-actions';
import {
  createBuildSetOverlayVisibleMainDepsHandler,
  createBuildToggleOverlayMainDepsHandler,
} from './overlay-main-actions-main-deps';

type SetVisibleOverlayVisibleMainDeps = Parameters<
  typeof createBuildSetVisibleOverlayVisibleMainDepsHandler
>[0];

export type OverlayVisibilityRuntimeDeps = {
  setVisibleOverlayVisibleDeps: SetVisibleOverlayVisibleMainDeps;
  getVisibleOverlayVisible: () => boolean;
};

export function createOverlayVisibilityRuntime(deps: OverlayVisibilityRuntimeDeps) {
  const setVisibleOverlayVisibleMainDeps = createBuildSetVisibleOverlayVisibleMainDepsHandler({
    ...deps.setVisibleOverlayVisibleDeps,
    getVisibleOverlayVisible: deps.getVisibleOverlayVisible,
  })();
  const setVisibleOverlayVisible = createSetVisibleOverlayVisibleHandler(
    setVisibleOverlayVisibleMainDeps,
  );

  const toggleVisibleOverlayMainDeps = createBuildToggleVisibleOverlayMainDepsHandler({
    getVisibleOverlayVisible: deps.getVisibleOverlayVisible,
    setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
  })();
  const toggleVisibleOverlay = createToggleVisibleOverlayHandler(toggleVisibleOverlayMainDeps);

  const setOverlayVisibleMainDeps = createBuildSetOverlayVisibleMainDepsHandler({
    setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
  })();
  const setOverlayVisible = createSetOverlayVisibleHandler(setOverlayVisibleMainDeps);

  const toggleOverlayMainDeps = createBuildToggleOverlayMainDepsHandler({
    toggleVisibleOverlay: () => toggleVisibleOverlay(),
  })();
  const toggleOverlay = createToggleOverlayHandler(toggleOverlayMainDeps);

  return {
    setVisibleOverlayVisible,
    toggleVisibleOverlay,
    setOverlayVisible,
    toggleOverlay,
  };
}
