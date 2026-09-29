type MpvCommand = (string | number)[];

function isSubtitleSeekCommand(command: MpvCommand): command is [string, number] {
  return command[0] === 'sub-seek' && typeof command[1] === 'number';
}

/**
 * Sends a configured mpv command. Subtitle seeks re-pause afterwards unless playback
 * is known to be running, so stepping lines from a paused video stays paused.
 */
export function dispatchConfiguredMpvCommand(
  command: MpvCommand,
  deps: {
    getPlaybackPaused: () => Promise<boolean | null>;
    sendMpvCommand: (command: MpvCommand) => void;
  },
): void {
  if (!isSubtitleSeekCommand(command)) {
    deps.sendMpvCommand(command);
    return;
  }

  void deps
    .getPlaybackPaused()
    .then((paused) => {
      deps.sendMpvCommand(command);
      if (paused !== false) {
        deps.sendMpvCommand(['set_property', 'pause', 'yes']);
      }
    })
    .catch(() => {
      deps.sendMpvCommand(command);
    });
}
