import {
  createBroadcastRuntimeOptionsChangedHandler,
  createOpenRuntimeOptionsPaletteHandler,
  createRestorePreviousSecondarySubVisibilityHandler,
  createSendToActiveOverlayWindowHandler,
  createSetOverlayDebugVisualizationEnabledHandler,
} from '../overlay-runtime-main-actions';
import type { ComposerInputs, ComposerOutputs } from './contracts';

type RestorePreviousSecondarySubVisibilityMainDeps = Parameters<
  typeof createRestorePreviousSecondarySubVisibilityHandler
>[0];
type BroadcastRuntimeOptionsChangedMainDeps = Parameters<
  typeof createBroadcastRuntimeOptionsChangedHandler
>[0];
type SendToActiveOverlayWindowMainDeps = Parameters<
  typeof createSendToActiveOverlayWindowHandler
>[0];
type SetOverlayDebugVisualizationEnabledMainDeps = Parameters<
  typeof createSetOverlayDebugVisualizationEnabledHandler
>[0];
type OpenRuntimeOptionsPaletteMainDeps = Parameters<
  typeof createOpenRuntimeOptionsPaletteHandler
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
    updateVisibleOverlayVisibility: () =>
      options.overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
    restorePreviousSecondarySubVisibility: createRestorePreviousSecondarySubVisibilityHandler(
      options.restorePreviousSecondarySubVisibilityMainDeps,
    ),
    broadcastRuntimeOptionsChanged: createBroadcastRuntimeOptionsChangedHandler(
      options.broadcastRuntimeOptionsChangedMainDeps,
    ),
    sendToActiveOverlayWindow: createSendToActiveOverlayWindowHandler(
      options.sendToActiveOverlayWindowMainDeps,
    ),
    setOverlayDebugVisualizationEnabled: createSetOverlayDebugVisualizationEnabledHandler(
      options.setOverlayDebugVisualizationEnabledMainDeps,
    ),
    openRuntimeOptionsPalette: createOpenRuntimeOptionsPaletteHandler(
      options.openRuntimeOptionsPaletteMainDeps,
    ),
  };
}
