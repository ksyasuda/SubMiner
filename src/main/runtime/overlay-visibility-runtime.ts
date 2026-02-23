import {
  createSetInvisibleOverlayVisibleHandler,
  createSetVisibleOverlayVisibleHandler,
  createToggleInvisibleOverlayHandler,
  createToggleVisibleOverlayHandler,
} from './overlay-visibility-actions';
import {
  createBuildSetInvisibleOverlayVisibleMainDepsHandler,
  createBuildSetVisibleOverlayVisibleMainDepsHandler,
  createBuildToggleInvisibleOverlayMainDepsHandler,
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
type SetInvisibleOverlayVisibleMainDeps = Parameters<
  typeof createBuildSetInvisibleOverlayVisibleMainDepsHandler
>[0];

export type OverlayVisibilityRuntimeDeps = {
  setVisibleOverlayVisibleDeps: SetVisibleOverlayVisibleMainDeps;
  setInvisibleOverlayVisibleDeps: SetInvisibleOverlayVisibleMainDeps;
  getVisibleOverlayVisible: () => boolean;
  getInvisibleOverlayVisible: () => boolean;
};

export function createOverlayVisibilityRuntime(deps: OverlayVisibilityRuntimeDeps) {
  const setVisibleOverlayVisibleMainDeps = createBuildSetVisibleOverlayVisibleMainDepsHandler(
    deps.setVisibleOverlayVisibleDeps,
  )();
  const setVisibleOverlayVisible = createSetVisibleOverlayVisibleHandler(
    setVisibleOverlayVisibleMainDeps,
  );

  const setInvisibleOverlayVisibleMainDeps = createBuildSetInvisibleOverlayVisibleMainDepsHandler(
    deps.setInvisibleOverlayVisibleDeps,
  )();
  const setInvisibleOverlayVisible = createSetInvisibleOverlayVisibleHandler(
    setInvisibleOverlayVisibleMainDeps,
  );

  const toggleVisibleOverlayMainDeps = createBuildToggleVisibleOverlayMainDepsHandler({
    getVisibleOverlayVisible: deps.getVisibleOverlayVisible,
    setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
  })();
  const toggleVisibleOverlay = createToggleVisibleOverlayHandler(toggleVisibleOverlayMainDeps);

  const toggleInvisibleOverlayMainDeps = createBuildToggleInvisibleOverlayMainDepsHandler({
    getInvisibleOverlayVisible: deps.getInvisibleOverlayVisible,
    setInvisibleOverlayVisible: (visible) => setInvisibleOverlayVisible(visible),
  })();
  const toggleInvisibleOverlay = createToggleInvisibleOverlayHandler(toggleInvisibleOverlayMainDeps);

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
    setInvisibleOverlayVisible,
    toggleVisibleOverlay,
    toggleInvisibleOverlay,
    setOverlayVisible,
    toggleOverlay,
  };
}
