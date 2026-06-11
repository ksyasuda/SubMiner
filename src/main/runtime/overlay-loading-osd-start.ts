export function shouldStartOverlayLoadingOsd(args: {
  visibleOverlayRequested: boolean;
  overlayContentReady: boolean;
  mediaPath?: string | null;
}): boolean {
  if (!args.visibleOverlayRequested || args.overlayContentReady) {
    return false;
  }
  if (args.mediaPath !== undefined && (args.mediaPath ?? '').trim().length === 0) {
    return false;
  }
  return true;
}

export function createMaybeStartOverlayLoadingOsdHandler(deps: {
  getVisibleOverlayRequested: () => boolean;
  isOverlayContentReady: () => boolean;
  startOverlayLoadingOsd: () => void;
}) {
  return (mediaPath?: string | null): void => {
    if (
      !shouldStartOverlayLoadingOsd({
        visibleOverlayRequested: deps.getVisibleOverlayRequested(),
        overlayContentReady: deps.isOverlayContentReady(),
        mediaPath,
      })
    ) {
      return;
    }
    deps.startOverlayLoadingOsd();
  };
}
