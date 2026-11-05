const OVERLAY_CONTROLLER_REFRESH_MS = 1_000;

/**
 * Decides whether dialogs show their controller shortcut tags. A fresh dialog window cannot
 * see a pad until a button is pressed in it, so an open dialog also asks main which controller
 * the overlay is reading. Called every frame; requests are limited to one per refresh interval.
 */
export function createControllerHintVisibility(options: {
  getOverlayControllerId: () => Promise<string | null>;
}) {
  let overlayHasController = false;
  let nextRefreshAt = Number.NEGATIVE_INFINITY;

  function refreshOverlayController(now: number): void {
    if (now < nextRefreshAt) return;
    nextRefreshAt = now + OVERLAY_CONTROLLER_REFRESH_MS;
    options.getOverlayControllerId().then(
      (gamepadId) => {
        overlayHasController = gamepadId !== null;
      },
      () => {
        // Keep the last answer; the next refresh tries again.
      },
    );
  }

  return {
    update(state: {
      enabled: boolean;
      dialogOpen: boolean;
      activeGamepadId: string | null;
      now: number;
    }): boolean {
      if (!state.enabled) return false;
      if (state.activeGamepadId !== null) return true;
      if (state.dialogOpen) refreshOverlayController(state.now);
      return overlayHasController;
    },
  };
}
