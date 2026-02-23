import type {
  createSetInvisibleOverlayVisibleHandler,
  createSetVisibleOverlayVisibleHandler,
  createToggleInvisibleOverlayHandler,
  createToggleVisibleOverlayHandler,
} from './overlay-visibility-actions';

type SetVisibleOverlayVisibleMainDeps = Parameters<typeof createSetVisibleOverlayVisibleHandler>[0];
type SetInvisibleOverlayVisibleMainDeps = Parameters<typeof createSetInvisibleOverlayVisibleHandler>[0];
type ToggleVisibleOverlayMainDeps = Parameters<typeof createToggleVisibleOverlayHandler>[0];
type ToggleInvisibleOverlayMainDeps = Parameters<typeof createToggleInvisibleOverlayHandler>[0];

export function createBuildSetVisibleOverlayVisibleMainDepsHandler(
  deps: SetVisibleOverlayVisibleMainDeps,
) {
  return (): SetVisibleOverlayVisibleMainDeps => ({
    setVisibleOverlayVisibleCore: (options) => deps.setVisibleOverlayVisibleCore(options),
    setVisibleOverlayVisibleState: (visible: boolean) => deps.setVisibleOverlayVisibleState(visible),
    updateVisibleOverlayVisibility: () => deps.updateVisibleOverlayVisibility(),
    updateInvisibleOverlayVisibility: () => deps.updateInvisibleOverlayVisibility(),
    syncInvisibleOverlayMousePassthrough: () => deps.syncInvisibleOverlayMousePassthrough(),
    shouldBindVisibleOverlayToMpvSubVisibility: () => deps.shouldBindVisibleOverlayToMpvSubVisibility(),
    isMpvConnected: () => deps.isMpvConnected(),
    setMpvSubVisibility: (visible: boolean) => deps.setMpvSubVisibility(visible),
  });
}

export function createBuildSetInvisibleOverlayVisibleMainDepsHandler(
  deps: SetInvisibleOverlayVisibleMainDeps,
) {
  return (): SetInvisibleOverlayVisibleMainDeps => ({
    setInvisibleOverlayVisibleCore: (options) => deps.setInvisibleOverlayVisibleCore(options),
    setInvisibleOverlayVisibleState: (visible: boolean) => deps.setInvisibleOverlayVisibleState(visible),
    updateInvisibleOverlayVisibility: () => deps.updateInvisibleOverlayVisibility(),
    syncInvisibleOverlayMousePassthrough: () => deps.syncInvisibleOverlayMousePassthrough(),
  });
}

export function createBuildToggleVisibleOverlayMainDepsHandler(deps: ToggleVisibleOverlayMainDeps) {
  return (): ToggleVisibleOverlayMainDeps => ({
    getVisibleOverlayVisible: () => deps.getVisibleOverlayVisible(),
    setVisibleOverlayVisible: (visible: boolean) => deps.setVisibleOverlayVisible(visible),
  });
}

export function createBuildToggleInvisibleOverlayMainDepsHandler(
  deps: ToggleInvisibleOverlayMainDeps,
) {
  return (): ToggleInvisibleOverlayMainDeps => ({
    getInvisibleOverlayVisible: () => deps.getInvisibleOverlayVisible(),
    setInvisibleOverlayVisible: (visible: boolean) => deps.setInvisibleOverlayVisible(visible),
  });
}
