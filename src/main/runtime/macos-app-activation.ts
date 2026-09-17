export type MacOSAppActivationDeps = {
  platform?: NodeJS.Platform;
  stealAppFocus: () => void;
  warn?: (message: string, details?: unknown) => void;
};

// macOS never promotes a background process to the foreground just because it opened a window:
// `BrowserWindow.show()`/`focus()` only reorder windows *within* the already-active app. Both ways
// `subminer anime` reaches the window (a cold launch from a terminal, or the command routed over
// the control socket to an instance that is already running behind mpv) leave another app
// frontmost, so the window appears buried. Activating the app itself is what pulls it forward.
//
// Callers must make sure the app is no longer a macOS accessory process first (see
// retainDockIcon) — an accessory app cannot become frontmost at all.
export function activateMacOSApp(deps: MacOSAppActivationDeps): boolean {
  if ((deps.platform ?? process.platform) !== 'darwin') {
    return false;
  }

  try {
    deps.stealAppFocus();
    return true;
  } catch (error) {
    deps.warn?.('Failed to activate app for foreground window', error);
    return false;
  }
}
