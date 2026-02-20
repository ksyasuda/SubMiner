type MpvClientCtorBaseOptions = {
  getResolvedConfig: () => unknown;
  autoStartOverlay: boolean;
  setOverlayVisible: (visible: boolean) => void;
  shouldBindVisibleOverlayToMpvSubVisibility: () => boolean;
  isVisibleOverlayVisible: () => boolean;
  getReconnectTimer: () => ReturnType<typeof setTimeout> | null;
  setReconnectTimer: (timer: ReturnType<typeof setTimeout> | null) => void;
};

type MpvClientLike = {
  connect: () => void;
};

type MpvClientCtor<TClient extends MpvClientLike, TOptions extends MpvClientCtorBaseOptions> = new (
  socketPath: string,
  options: TOptions,
) => TClient;

export function createMpvClientRuntimeServiceFactory<
  TClient extends MpvClientLike,
  TOptions extends MpvClientCtorBaseOptions,
>(deps: {
  createClient: MpvClientCtor<TClient, TOptions>;
  socketPath: string;
  options: TOptions;
  bindEventHandlers: (client: TClient) => void;
}) {
  return (): TClient => {
    const mpvClient = new deps.createClient(deps.socketPath, deps.options);
    deps.bindEventHandlers(mpvClient);
    mpvClient.connect();
    return mpvClient;
  };
}
