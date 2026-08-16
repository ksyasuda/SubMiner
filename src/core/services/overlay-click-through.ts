type ClickThroughWindow = {
  setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => void;
};

/**
 * Puts an overlay window into click-through mode. Forwarded mouse-move ({ forward: true }) is
 * what lets renderer hover tracking wake a click-through overlay, but on Windows Electron
 * implements it with a global WH_MOUSE_LL hook whose callback runs on the main-process message
 * loop, so any main-thread stall delays mouse input system-wide (electron/electron#10183).
 * Windows instead wakes the overlay via the main-process cursor poll
 * (tickWindowsOverlayPointerInteraction), so no forwarding is requested there. macOS still
 * needs forwarding for renderer hover tracking; Linux ignores the flag entirely
 * (electron/electron#16777).
 *
 * Pass isWindowsPlatform when the caller already carries a platform flag (tests simulate
 * platforms through it); otherwise the real process.platform decides.
 */
export function applyOverlayClickThrough(
  window: ClickThroughWindow,
  isWindowsPlatform?: boolean,
): void {
  if (isWindowsPlatform ?? process.platform === 'win32') {
    window.setIgnoreMouseEvents(true);
  } else {
    window.setIgnoreMouseEvents(true, { forward: true });
  }
}
