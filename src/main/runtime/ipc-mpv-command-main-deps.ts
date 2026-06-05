import type { MpvCommandFromIpcRuntimeDeps } from '../ipc-mpv-command';

export function createBuildMpvCommandFromIpcRuntimeMainDepsHandler(
  deps: MpvCommandFromIpcRuntimeDeps,
) {
  return (): MpvCommandFromIpcRuntimeDeps => {
    const showPlaybackFeedback = deps.showPlaybackFeedback;
    return {
      triggerSubsyncFromConfig: () => deps.triggerSubsyncFromConfig(),
      openRuntimeOptionsPalette: () => deps.openRuntimeOptionsPalette(),
      openJimaku: () => deps.openJimaku(),
      openYoutubeTrackPicker: () => deps.openYoutubeTrackPicker(),
      openPlaylistBrowser: () => deps.openPlaylistBrowser(),
      cycleRuntimeOption: (id, direction) => deps.cycleRuntimeOption(id, direction),
      showMpvOsd: (text: string) => deps.showMpvOsd(text),
      ...(showPlaybackFeedback
        ? { showPlaybackFeedback: (text: string) => showPlaybackFeedback(text) }
        : {}),
      replayCurrentSubtitle: () => deps.replayCurrentSubtitle(),
      playNextSubtitle: () => deps.playNextSubtitle(),
      shiftSubDelayToAdjacentSubtitle: (direction) =>
        deps.shiftSubDelayToAdjacentSubtitle(direction),
      sendMpvCommand: (command: (string | number)[]) => deps.sendMpvCommand(command),
      getMpvClient: () => deps.getMpvClient(),
      isMpvConnected: () => deps.isMpvConnected(),
      hasRuntimeOptionsManager: () => deps.hasRuntimeOptionsManager(),
    };
  };
}
