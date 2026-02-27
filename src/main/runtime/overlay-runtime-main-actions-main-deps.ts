import {
  createBroadcastRuntimeOptionsChangedHandler,
  createGetRuntimeOptionsStateHandler,
  createOpenRuntimeOptionsPaletteHandler,
  createRestorePreviousSecondarySubVisibilityHandler,
  createSendToActiveOverlayWindowHandler,
  createSetOverlayDebugVisualizationEnabledHandler,
} from './overlay-runtime-main-actions';

type GetRuntimeOptionsStateMainDeps = Parameters<typeof createGetRuntimeOptionsStateHandler>[0];
type RestorePreviousSecondarySubVisibilityMainDeps = Parameters<
  typeof createRestorePreviousSecondarySubVisibilityHandler
>[0];
type BroadcastRuntimeOptionsChangedMainDeps = Parameters<
  typeof createBroadcastRuntimeOptionsChangedHandler
>[0];
type SendToActiveOverlayWindowMainDeps = Parameters<typeof createSendToActiveOverlayWindowHandler>[0];
type SetOverlayDebugVisualizationEnabledMainDeps = Parameters<
  typeof createSetOverlayDebugVisualizationEnabledHandler
>[0];
type OpenRuntimeOptionsPaletteMainDeps = Parameters<typeof createOpenRuntimeOptionsPaletteHandler>[0];

export function createBuildGetRuntimeOptionsStateMainDepsHandler(
  deps: GetRuntimeOptionsStateMainDeps,
) {
  return (): GetRuntimeOptionsStateMainDeps => ({
    getRuntimeOptionsManager: () => deps.getRuntimeOptionsManager(),
  });
}

export function createBuildRestorePreviousSecondarySubVisibilityMainDepsHandler(
  deps: RestorePreviousSecondarySubVisibilityMainDeps,
) {
  return (): RestorePreviousSecondarySubVisibilityMainDeps => ({
    getMpvClient: () => deps.getMpvClient(),
  });
}

export function createBuildBroadcastRuntimeOptionsChangedMainDepsHandler(
  deps: BroadcastRuntimeOptionsChangedMainDeps,
) {
  return (): BroadcastRuntimeOptionsChangedMainDeps => ({
    broadcastRuntimeOptionsChangedRuntime: (getRuntimeOptionsState, broadcastToOverlayWindows) =>
      deps.broadcastRuntimeOptionsChangedRuntime(getRuntimeOptionsState, broadcastToOverlayWindows),
    getRuntimeOptionsState: () => deps.getRuntimeOptionsState(),
    broadcastToOverlayWindows: (channel: string, ...args: unknown[]) =>
      deps.broadcastToOverlayWindows(channel, ...args),
  });
}

export function createBuildSendToActiveOverlayWindowMainDepsHandler(
  deps: SendToActiveOverlayWindowMainDeps,
) {
  return (): SendToActiveOverlayWindowMainDeps => ({
    sendToActiveOverlayWindowRuntime: (channel, payload, runtimeOptions) =>
      deps.sendToActiveOverlayWindowRuntime(channel, payload, runtimeOptions),
  });
}

export function createBuildSetOverlayDebugVisualizationEnabledMainDepsHandler(
  deps: SetOverlayDebugVisualizationEnabledMainDeps,
) {
  return (): SetOverlayDebugVisualizationEnabledMainDeps => ({
    setOverlayDebugVisualizationEnabledRuntime: (
      currentEnabled,
      nextEnabled,
      setCurrentEnabled,
    ) =>
      deps.setOverlayDebugVisualizationEnabledRuntime(
        currentEnabled,
        nextEnabled,
        setCurrentEnabled,
      ),
    getCurrentEnabled: () => deps.getCurrentEnabled(),
    setCurrentEnabled: (enabled: boolean) => deps.setCurrentEnabled(enabled),
  });
}

export function createBuildOpenRuntimeOptionsPaletteMainDepsHandler(
  deps: OpenRuntimeOptionsPaletteMainDeps,
) {
  return (): OpenRuntimeOptionsPaletteMainDeps => ({
    openRuntimeOptionsPaletteRuntime: () => deps.openRuntimeOptionsPaletteRuntime(),
  });
}
