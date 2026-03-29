import {
  createBroadcastRuntimeOptionsChangedHandler,
  createOpenRuntimeOptionsPaletteHandler,
  createRestorePreviousSecondarySubVisibilityHandler,
  createSendToActiveOverlayWindowHandler,
  createSetOverlayDebugVisualizationEnabledHandler,
} from '../overlay-runtime-main-actions';
import {
  createBuildBroadcastRuntimeOptionsChangedMainDepsHandler,
  createBuildOpenRuntimeOptionsPaletteMainDepsHandler,
  createBuildRestorePreviousSecondarySubVisibilityMainDepsHandler,
  createBuildSendToActiveOverlayWindowMainDepsHandler,
  createBuildSetOverlayDebugVisualizationEnabledMainDepsHandler,
} from '../overlay-runtime-main-actions-main-deps';
import type { ComposerInputs, ComposerOutputs } from './contracts';

type RestorePreviousSecondarySubVisibilityMainDeps = Parameters<
  typeof createBuildRestorePreviousSecondarySubVisibilityMainDepsHandler
>[0];
type BroadcastRuntimeOptionsChangedMainDeps = Parameters<
  typeof createBuildBroadcastRuntimeOptionsChangedMainDepsHandler
>[0];
type SendToActiveOverlayWindowMainDeps = Parameters<
  typeof createBuildSendToActiveOverlayWindowMainDepsHandler
>[0];
type SetOverlayDebugVisualizationEnabledMainDeps = Parameters<
  typeof createBuildSetOverlayDebugVisualizationEnabledMainDepsHandler
>[0];
type OpenRuntimeOptionsPaletteMainDeps = Parameters<
  typeof createBuildOpenRuntimeOptionsPaletteMainDepsHandler
>[0];

export type OverlayVisibilityRuntimeComposerOptions = ComposerInputs<{
  overlayVisibilityRuntime: {
    updateVisibleOverlayVisibility: () => void;
  };
  restorePreviousSecondarySubVisibilityMainDeps: RestorePreviousSecondarySubVisibilityMainDeps;
  broadcastRuntimeOptionsChangedMainDeps: BroadcastRuntimeOptionsChangedMainDeps;
  sendToActiveOverlayWindowMainDeps: SendToActiveOverlayWindowMainDeps;
  setOverlayDebugVisualizationEnabledMainDeps: SetOverlayDebugVisualizationEnabledMainDeps;
  openRuntimeOptionsPaletteMainDeps: OpenRuntimeOptionsPaletteMainDeps;
}>;

export type OverlayVisibilityRuntimeComposerResult = ComposerOutputs<{
  updateVisibleOverlayVisibility: () => void;
  restorePreviousSecondarySubVisibility: ReturnType<
    typeof createRestorePreviousSecondarySubVisibilityHandler
  >;
  broadcastRuntimeOptionsChanged: ReturnType<typeof createBroadcastRuntimeOptionsChangedHandler>;
  sendToActiveOverlayWindow: ReturnType<typeof createSendToActiveOverlayWindowHandler>;
  setOverlayDebugVisualizationEnabled: ReturnType<
    typeof createSetOverlayDebugVisualizationEnabledHandler
  >;
  openRuntimeOptionsPalette: ReturnType<typeof createOpenRuntimeOptionsPaletteHandler>;
}>;

export function composeOverlayVisibilityRuntime(
  options: OverlayVisibilityRuntimeComposerOptions,
): OverlayVisibilityRuntimeComposerResult {
  return {
    updateVisibleOverlayVisibility: () => options.overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
    restorePreviousSecondarySubVisibility: createRestorePreviousSecondarySubVisibilityHandler(
      createBuildRestorePreviousSecondarySubVisibilityMainDepsHandler(
        options.restorePreviousSecondarySubVisibilityMainDeps,
      )(),
    ),
    broadcastRuntimeOptionsChanged: createBroadcastRuntimeOptionsChangedHandler(
      createBuildBroadcastRuntimeOptionsChangedMainDepsHandler(
        options.broadcastRuntimeOptionsChangedMainDeps,
      )(),
    ),
    sendToActiveOverlayWindow: createSendToActiveOverlayWindowHandler(
      createBuildSendToActiveOverlayWindowMainDepsHandler(
        options.sendToActiveOverlayWindowMainDeps,
      )(),
    ),
    setOverlayDebugVisualizationEnabled: createSetOverlayDebugVisualizationEnabledHandler(
      createBuildSetOverlayDebugVisualizationEnabledMainDepsHandler(
        options.setOverlayDebugVisualizationEnabledMainDeps,
      )(),
    ),
    openRuntimeOptionsPalette: createOpenRuntimeOptionsPaletteHandler(
      createBuildOpenRuntimeOptionsPaletteMainDepsHandler(
        options.openRuntimeOptionsPaletteMainDeps,
      )(),
    ),
  };
}
