type SeparateWindowLike = {
  isDestroyed(): boolean;
  isVisible?: () => boolean;
};

export function hasLiveSeparateWindow(
  windows: Array<SeparateWindowLike | null | undefined>,
): boolean {
  return windows.some(
    (window) =>
      Boolean(window && !window.isDestroyed()) &&
      (typeof window?.isVisible !== 'function' || window.isVisible()),
  );
}

export function shouldSuppressVisibleOverlayRaiseForSeparateWindow(options: {
  window: unknown;
  mainWindow: unknown;
  separateWindows: Array<SeparateWindowLike | null | undefined>;
}): boolean {
  if (!options.mainWindow || options.window !== options.mainWindow) {
    return false;
  }

  return hasLiveSeparateWindow(options.separateWindows);
}
