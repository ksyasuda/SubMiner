import type { MpvCommandFromIpcRuntimeDeps } from '../ipc-mpv-command';

export function createBuildMpvCommandFromIpcRuntimeMainDepsHandler(
  deps: MpvCommandFromIpcRuntimeDeps,
) {
  return (): MpvCommandFromIpcRuntimeDeps => {
    const showPlaybackFeedback = deps.showPlaybackFeedback;
    const showRawMpvOsd = deps.showRawMpvOsd;
    return {
      triggerSubsyncFromConfig: () => deps.triggerSubsyncFromConfig(),
      openRuntimeOptionsPalette: () => deps.openRuntimeOptionsPalette(),
      openJimaku: () => deps.openJimaku(),
      openAnimetosho: () => deps.openAnimetosho(),
      openYoutubeTrackPicker: () => deps.openYoutubeTrackPicker(),
      openPlaylistBrowser: () => deps.openPlaylistBrowser(),
      cycleRuntimeOption: (id, direction) => deps.cycleRuntimeOption(id, direction),
      showMpvOsd: (text: string) => deps.showMpvOsd(text),
      ...(showRawMpvOsd ? { showRawMpvOsd: (text: string) => showRawMpvOsd(text) } : {}),
      ...(showPlaybackFeedback
        ? { showPlaybackFeedback: (text: string) => showPlaybackFeedback(text) }
        : {}),
      replayCurrentSubtitle: () => deps.replayCurrentSubtitle(),
      playNextSubtitle: () => deps.playNextSubtitle(),
      sendMpvCommand: (command: (string | number)[]) => deps.sendMpvCommand(command),
      getMpvClient: () => deps.getMpvClient(),
      isMpvConnected: () => deps.isMpvConnected(),
      hasRuntimeOptionsManager: () => deps.hasRuntimeOptionsManager(),
    };
  };
}
