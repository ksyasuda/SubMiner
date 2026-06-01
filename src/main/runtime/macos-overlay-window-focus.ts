type FocusableOverlayWebContents = {
  isFocused: () => boolean;
  focus: () => void;
};

type FocusableOverlayWindow = {
  isDestroyed: () => boolean;
  isVisible: () => boolean;
  isFocused: () => boolean;
  setFocusable?: (focusable: boolean) => void;
  focus: () => void;
  webContents: FocusableOverlayWebContents;
};

export type MacOSOverlayWindowFocusDeps = {
  platform: NodeJS.Platform;
  getOverlayWindow: () => FocusableOverlayWindow | null;
  stealAppFocus: () => void;
  warn: (message: string, details?: unknown) => void;
};

// macOS only delivers mouse-moved/hover events to the key window of the frontmost application.
// After autoplay warmup completes mpv is the frontmost process, so the transparent overlay window
// receives no pointer events until the user physically clicks a subtitle (which activates the app
// via acceptFirstMouse). Renderer-side pointer recovery can toggle setIgnoreMouseEvents but cannot
// make the window key, so it cannot wake hover on its own. Activating the overlay window from the
// main process reproduces that manual click and keeps subtitles interactive.
// (Modal close takes the opposite path — see restoreMacOSMpvFocusAfterModalClose — because the user
// needs keyboard focus back on mpv, with the overlay floating passively above it.)
export function focusMacOSOverlayWindow(deps: MacOSOverlayWindowFocusDeps): void {
  if (deps.platform !== 'darwin') {
    return;
  }

  const overlayWindow = deps.getOverlayWindow();
  if (!overlayWindow || overlayWindow.isDestroyed() || !overlayWindow.isVisible()) {
    return;
  }
  if (overlayWindow.isFocused()) {
    return;
  }

  try {
    deps.stealAppFocus();
  } catch (error) {
    deps.warn('Failed to steal app focus for overlay window', error);
  }

  overlayWindow.setFocusable?.(true);
  overlayWindow.focus();
  if (!overlayWindow.webContents.isFocused()) {
    overlayWindow.webContents.focus();
  }
}
