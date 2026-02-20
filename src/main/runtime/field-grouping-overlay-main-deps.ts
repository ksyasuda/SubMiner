export function createBuildFieldGroupingOverlayMainDepsHandler<
  TModal extends string,
  TChoice,
>(deps: {
  getMainWindow: () => unknown | null;
  getVisibleOverlayVisible: () => boolean;
  getInvisibleOverlayVisible: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  setInvisibleOverlayVisible: (visible: boolean) => void;
  getResolver: () => ((choice: TChoice) => void) | null;
  setResolver: (resolver: ((choice: TChoice) => void) | null) => void;
  getRestoreVisibleOverlayOnModalClose: () => Set<TModal>;
  sendToActiveOverlayWindow: (
    channel: string,
    payload?: unknown,
    runtimeOptions?: { restoreOnModalClose?: TModal },
  ) => boolean;
}) {
  return () => ({
    getMainWindow: () => deps.getMainWindow() as never,
    getVisibleOverlayVisible: () => deps.getVisibleOverlayVisible(),
    getInvisibleOverlayVisible: () => deps.getInvisibleOverlayVisible(),
    setVisibleOverlayVisible: (visible: boolean) => deps.setVisibleOverlayVisible(visible),
    setInvisibleOverlayVisible: (visible: boolean) => deps.setInvisibleOverlayVisible(visible),
    getResolver: () => deps.getResolver() as never,
    setResolver: (resolver: ((choice: TChoice) => void) | null) => deps.setResolver(resolver),
    getRestoreVisibleOverlayOnModalClose: () => deps.getRestoreVisibleOverlayOnModalClose(),
    sendToVisibleOverlay: (
      channel: string,
      payload?: unknown,
      runtimeOptions?: { restoreOnModalClose?: TModal },
    ) => deps.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
  });
}
