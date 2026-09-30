type MpvCommand = (string | number)[];

function isSubtitleSeekCommand(command: MpvCommand): command is [string, number] {
  return command[0] === 'sub-seek' && typeof command[1] === 'number';
}

/**
 * Sends a configured mpv command. Subtitle seeks re-pause afterwards unless playback
 * is known to be running, so stepping lines from a paused video stays paused.
 * Resolves once every command has been sent; it never rejects.
 */
export function dispatchConfiguredMpvCommand(
  command: MpvCommand,
  deps: {
    getPlaybackPaused: () => Promise<boolean | null>;
    sendMpvCommand: (command: MpvCommand) => void;
  },
): Promise<void> {
  if (!isSubtitleSeekCommand(command)) {
    deps.sendMpvCommand(command);
    return Promise.resolve();
  }

  // The fallback only covers a failed pause lookup, so a failed re-pause never resends the seek.
  return deps
    .getPlaybackPaused()
    .then(
      (paused) => {
        deps.sendMpvCommand(command);
        if (paused !== false) {
          deps.sendMpvCommand(['set_property', 'pause', 'yes']);
        }
      },
      () => {
        deps.sendMpvCommand(command);
      },
    )
    .catch((error: unknown) => console.error('Could not send mpv command', error));
}
