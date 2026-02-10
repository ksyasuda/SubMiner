export function syncInvisibleOverlayMousePassthroughService(options: {
  hasInvisibleWindow: () => boolean;
  setIgnoreMouseEvents: (ignore: boolean, extra?: { forward: boolean }) => void;
  visibleOverlayVisible: boolean;
  invisibleOverlayVisible: boolean;
}): void {
  if (!options.hasInvisibleWindow()) return;
  if (options.visibleOverlayVisible) {
    options.setIgnoreMouseEvents(true, { forward: true });
  } else if (options.invisibleOverlayVisible) {
    options.setIgnoreMouseEvents(false);
  }
}

export function setVisibleOverlayVisibleService(options: {
  visible: boolean;
  setVisibleOverlayVisibleState: (visible: boolean) => void;
  updateVisibleOverlayVisibility: () => void;
  updateInvisibleOverlayVisibility: () => void;
  syncInvisibleOverlayMousePassthrough: () => void;
  shouldBindVisibleOverlayToMpvSubVisibility: () => boolean;
  isMpvConnected: () => boolean;
  setMpvSubVisibility: (visible: boolean) => void;
}): void {
  options.setVisibleOverlayVisibleState(options.visible);
  options.updateVisibleOverlayVisibility();
  options.updateInvisibleOverlayVisibility();
  options.syncInvisibleOverlayMousePassthrough();
  if (
    options.shouldBindVisibleOverlayToMpvSubVisibility() &&
    options.isMpvConnected()
  ) {
    options.setMpvSubVisibility(!options.visible);
  }
}

export function setInvisibleOverlayVisibleService(options: {
  visible: boolean;
  setInvisibleOverlayVisibleState: (visible: boolean) => void;
  updateInvisibleOverlayVisibility: () => void;
  syncInvisibleOverlayMousePassthrough: () => void;
}): void {
  options.setInvisibleOverlayVisibleState(options.visible);
  options.updateInvisibleOverlayVisibility();
  options.syncInvisibleOverlayMousePassthrough();
}
