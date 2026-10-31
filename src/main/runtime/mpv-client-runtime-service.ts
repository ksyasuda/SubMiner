import type { Config } from '../../types';

export type MpvClientRuntimeServiceOptions = {
  getResolvedConfig: () => Config;
  autoStartOverlay: boolean;
  setOverlayVisible: (visible: boolean) => void;
  isVisibleOverlayVisible: () => boolean;
  getReconnectTimer: () => ReturnType<typeof setTimeout> | null;
  setReconnectTimer: (timer: ReturnType<typeof setTimeout> | null) => void;
  shouldAutoLoadSecondarySubTrack?: (path: string) => boolean;
  shouldQuitOnMpvShutdown?: () => boolean;
  requestAppQuit?: () => void;
};

type MpvClientLike = {
  connect: () => void;
};

// Creates, binds and connects a new mpv client. Socket path and auto-start state are read
// on each call, so a reconnect picks up the current values.
export function createMpvClientRuntimeServiceFactory<TClient extends MpvClientLike>(deps: {
  createClient: new (socketPath: string, options: MpvClientRuntimeServiceOptions) => TClient;
  getSocketPath: () => string;
  getResolvedConfig: () => Config;
  isAutoStartOverlayEnabled: () => boolean;
  setOverlayVisible: (visible: boolean) => void;
  isVisibleOverlayVisible: () => boolean;
  getReconnectTimer: () => ReturnType<typeof setTimeout> | null;
  setReconnectTimer: (timer: ReturnType<typeof setTimeout> | null) => void;
  shouldAutoLoadSecondarySubTrack?: (path: string) => boolean;
  shouldQuitOnMpvShutdown?: () => boolean;
  requestAppQuit?: () => void;
  bindEventHandlers: (client: TClient) => void;
}) {
  return (): TClient => {
    const socketPath = deps.getSocketPath();
    const mpvClient = new deps.createClient(socketPath, {
      getResolvedConfig: () => deps.getResolvedConfig(),
      autoStartOverlay: deps.isAutoStartOverlayEnabled(),
      setOverlayVisible: (visible) => deps.setOverlayVisible(visible),
      isVisibleOverlayVisible: () => deps.isVisibleOverlayVisible(),
      getReconnectTimer: () => deps.getReconnectTimer(),
      setReconnectTimer: (timer) => deps.setReconnectTimer(timer),
      shouldAutoLoadSecondarySubTrack: deps.shouldAutoLoadSecondarySubTrack,
      shouldQuitOnMpvShutdown: () => deps.shouldQuitOnMpvShutdown?.() ?? false,
      requestAppQuit: () => deps.requestAppQuit?.(),
    });
    deps.bindEventHandlers(mpvClient);
    mpvClient.connect();
    return mpvClient;
  };
}
