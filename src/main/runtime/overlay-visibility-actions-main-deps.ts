import type {
  createSetVisibleOverlayVisibleHandler,
  createToggleVisibleOverlayHandler,
} from './overlay-visibility-actions';

type SetVisibleOverlayVisibleMainDeps = Parameters<typeof createSetVisibleOverlayVisibleHandler>[0];
type ToggleVisibleOverlayMainDeps = Parameters<typeof createToggleVisibleOverlayHandler>[0];

export function createBuildSetVisibleOverlayVisibleMainDepsHandler(
  deps: SetVisibleOverlayVisibleMainDeps,
) {
  return (): SetVisibleOverlayVisibleMainDeps => ({
    setVisibleOverlayVisibleCore: (options) => deps.setVisibleOverlayVisibleCore(options),
    setVisibleOverlayVisibleState: (visible: boolean) => deps.setVisibleOverlayVisibleState(visible),
    updateVisibleOverlayVisibility: () => deps.updateVisibleOverlayVisibility(),
  });
}

export function createBuildToggleVisibleOverlayMainDepsHandler(deps: ToggleVisibleOverlayMainDeps) {
  return (): ToggleVisibleOverlayMainDeps => ({
    getVisibleOverlayVisible: () => deps.getVisibleOverlayVisible(),
    setVisibleOverlayVisible: (visible: boolean) => deps.setVisibleOverlayVisible(visible),
  });
}
