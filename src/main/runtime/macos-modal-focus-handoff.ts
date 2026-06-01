type RefreshableWindowTracker = {
  refreshNow: () => Promise<void>;
};

export type MacOSModalFocusHandoffDeps = {
  platform: NodeJS.Platform;
  focusMpv: () => Promise<void>;
  getWindowTracker: () => RefreshableWindowTracker | null;
  updateVisibleOverlayVisibility: () => void;
  warn: (message: string, details?: unknown) => void;
};

export async function restoreMacOSMpvFocusAfterModalClose(
  deps: MacOSModalFocusHandoffDeps,
): Promise<void> {
  if (deps.platform !== 'darwin') {
    return;
  }

  try {
    await deps.focusMpv();
  } catch (error) {
    deps.warn('Failed to focus mpv after macOS modal close', error);
  }

  try {
    await deps.getWindowTracker()?.refreshNow();
  } catch (error) {
    deps.warn('Failed to refresh macOS mpv focus after modal close', error);
  }

  deps.updateVisibleOverlayVisibility();
}
