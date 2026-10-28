type MpvCommand = Array<string | number>;

/**
 * Optional pause while keyboard-only mode is on (`subtitleStyle.autoPauseVideoOnKeyboardMode`).
 *
 * Entering the mode pauses a playing video. Leaving it resumes only a pause the mode made
 * itself: if the video was already paused, or the user played or paused it while in the
 * mode, leaving does nothing. The user taking over is seen through mpv's pause changes.
 */
export function createKeyboardModeAutoPause(deps: {
  isEnabled: () => boolean;
  getPlaybackPaused: () => Promise<boolean | null>;
  sendMpvCommand: (command: MpvCommand) => void;
}) {
  let pausedByMode = false;
  /** Our own pause command is still to be reported back by mpv. */
  let awaitingOwnPause = false;
  /** Bumped on every mode change so a slow pause lookup cannot act on a stale mode. */
  let generation = 0;

  async function enterMode(current: number): Promise<void> {
    if (!deps.isEnabled()) return;
    let paused: boolean | null = null;
    try {
      paused = await deps.getPlaybackPaused();
    } catch {
      return;
    }
    // Left the mode again (or re-entered) while waiting: this request is stale.
    if (current !== generation || paused !== false) return;
    pausedByMode = true;
    awaitingOwnPause = true;
    deps.sendMpvCommand(['set_property', 'pause', 'yes']);
  }

  function leaveMode(): void {
    if (!pausedByMode) return;
    pausedByMode = false;
    awaitingOwnPause = false;
    deps.sendMpvCommand(['set_property', 'pause', 'no']);
  }

  return {
    /** Keyboard-only mode turned on or off, from the keyboard or the controller. */
    handleModeChanged(enabled: boolean): Promise<void> {
      generation += 1;
      if (enabled) return enterMode(generation);
      leaveMode();
      return Promise.resolve();
    },

    /** mpv's pause property changed, from any source. */
    handlePlaybackPauseChanged(paused: boolean): void {
      if (!pausedByMode) return;
      if (paused && awaitingOwnPause) {
        awaitingOwnPause = false;
        return;
      }
      // Anything else (a resume, or a pause after one) is the user's own playback choice.
      pausedByMode = false;
      awaitingOwnPause = false;
    },

    isPausedByMode: (): boolean => pausedByMode,
  };
}
