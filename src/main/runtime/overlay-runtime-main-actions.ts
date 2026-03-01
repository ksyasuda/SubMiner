import type { RuntimeOptionState } from '../../types';
import type { OverlayHostedModal } from '../overlay-runtime';

type RuntimeOptionsManagerLike = {
  listOptions: () => RuntimeOptionState[];
};

type MpvClientLike = {
  connected: boolean;
  restorePreviousSecondarySubVisibility: () => void;
};

export function createGetRuntimeOptionsStateHandler(deps: {
  getRuntimeOptionsManager: () => RuntimeOptionsManagerLike | null;
}) {
  return (): RuntimeOptionState[] => {
    const manager = deps.getRuntimeOptionsManager();
    if (!manager) return [];
    return manager.listOptions();
  };
}

export function createRestorePreviousSecondarySubVisibilityHandler(deps: {
  getMpvClient: () => MpvClientLike | null;
}) {
  return (): void => {
    const client = deps.getMpvClient();
    if (!client || !client.connected) return;
    client.restorePreviousSecondarySubVisibility();
  };
}

export function createBroadcastRuntimeOptionsChangedHandler(deps: {
  broadcastRuntimeOptionsChangedRuntime: (
    getRuntimeOptionsState: () => RuntimeOptionState[],
    broadcastToOverlayWindows: (channel: string, ...args: unknown[]) => void,
  ) => void;
  getRuntimeOptionsState: () => RuntimeOptionState[];
  broadcastToOverlayWindows: (channel: string, ...args: unknown[]) => void;
}) {
  return (): void => {
    deps.broadcastRuntimeOptionsChangedRuntime(
      () => deps.getRuntimeOptionsState(),
      (channel, ...args) => deps.broadcastToOverlayWindows(channel, ...args),
    );
  };
}

export function createSendToActiveOverlayWindowHandler(deps: {
  sendToActiveOverlayWindowRuntime: (
    channel: string,
    payload?: unknown,
    runtimeOptions?: { restoreOnModalClose?: OverlayHostedModal },
  ) => boolean;
}) {
  return (
    channel: string,
    payload?: unknown,
    runtimeOptions?: { restoreOnModalClose?: OverlayHostedModal },
  ): boolean => deps.sendToActiveOverlayWindowRuntime(channel, payload, runtimeOptions);
}

export function createSetOverlayDebugVisualizationEnabledHandler(deps: {
  setOverlayDebugVisualizationEnabledRuntime: (
    currentEnabled: boolean,
    nextEnabled: boolean,
    setCurrentEnabled: (enabled: boolean) => void,
  ) => void;
  getCurrentEnabled: () => boolean;
  setCurrentEnabled: (enabled: boolean) => void;
}) {
  return (enabled: boolean): void => {
    deps.setOverlayDebugVisualizationEnabledRuntime(deps.getCurrentEnabled(), enabled, (next) =>
      deps.setCurrentEnabled(next),
    );
  };
}

export function createOpenRuntimeOptionsPaletteHandler(deps: {
  openRuntimeOptionsPaletteRuntime: () => void;
}) {
  return (): void => {
    deps.openRuntimeOptionsPaletteRuntime();
  };
}
