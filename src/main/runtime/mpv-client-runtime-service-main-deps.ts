export function createBuildMpvClientRuntimeServiceFactoryDepsHandler<
  TClient,
  TResolvedConfig,
  TOptions,
>(deps: {
  createClient: new (socketPath: string, options: TOptions) => TClient;
  getSocketPath: () => string;
  getResolvedConfig: () => TResolvedConfig;
  isAutoStartOverlayEnabled: () => boolean;
  setOverlayVisible: (visible: boolean) => void;
  isVisibleOverlayVisible: () => boolean;
  getReconnectTimer: () => ReturnType<typeof setTimeout> | null;
  setReconnectTimer: (timer: ReturnType<typeof setTimeout> | null) => void;
  shouldQuitOnMpvShutdown?: () => boolean;
  requestAppQuit?: () => void;
  bindEventHandlers: (client: TClient) => void;
}) {
  return () => ({
    createClient: deps.createClient,
    socketPath: deps.getSocketPath(),
    options: {
      getResolvedConfig: () => deps.getResolvedConfig(),
      autoStartOverlay: deps.isAutoStartOverlayEnabled(),
      setOverlayVisible: (visible: boolean) => deps.setOverlayVisible(visible),
      isVisibleOverlayVisible: () => deps.isVisibleOverlayVisible(),
      getReconnectTimer: () => deps.getReconnectTimer(),
      setReconnectTimer: (timer: ReturnType<typeof setTimeout> | null) =>
        deps.setReconnectTimer(timer),
      shouldQuitOnMpvShutdown: () => deps.shouldQuitOnMpvShutdown?.() ?? false,
      requestAppQuit: () => deps.requestAppQuit?.(),
    },
    bindEventHandlers: (client: TClient) => deps.bindEventHandlers(client),
  });
}
