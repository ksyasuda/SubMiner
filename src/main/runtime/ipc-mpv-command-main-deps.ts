import type { MpvCommandFromIpcRuntimeDeps } from '../ipc-mpv-command';

export function createBuildMpvCommandFromIpcRuntimeMainDepsHandler(
  deps: MpvCommandFromIpcRuntimeDeps,
) {
  return (): MpvCommandFromIpcRuntimeDeps => ({
    triggerSubsyncFromConfig: () => deps.triggerSubsyncFromConfig(),
    openRuntimeOptionsPalette: () => deps.openRuntimeOptionsPalette(),
    openYoutubeTrackPicker: () => deps.openYoutubeTrackPicker(),
    cycleRuntimeOption: (id, direction) => deps.cycleRuntimeOption(id, direction),
    showMpvOsd: (text: string) => deps.showMpvOsd(text),
    replayCurrentSubtitle: () => deps.replayCurrentSubtitle(),
    playNextSubtitle: () => deps.playNextSubtitle(),
    shiftSubDelayToAdjacentSubtitle: (direction) => deps.shiftSubDelayToAdjacentSubtitle(direction),
    sendMpvCommand: (command: (string | number)[]) => deps.sendMpvCommand(command),
    getMpvClient: () => deps.getMpvClient(),
    isMpvConnected: () => deps.isMpvConnected(),
    hasRuntimeOptionsManager: () => deps.hasRuntimeOptionsManager(),
  });
}
