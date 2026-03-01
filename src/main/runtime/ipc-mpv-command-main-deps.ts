import type { MpvCommandFromIpcRuntimeDeps } from '../ipc-mpv-command';

export function createBuildMpvCommandFromIpcRuntimeMainDepsHandler(
  deps: MpvCommandFromIpcRuntimeDeps,
) {
  return (): MpvCommandFromIpcRuntimeDeps => ({
    triggerSubsyncFromConfig: () => deps.triggerSubsyncFromConfig(),
    openRuntimeOptionsPalette: () => deps.openRuntimeOptionsPalette(),
    cycleRuntimeOption: (id, direction) => deps.cycleRuntimeOption(id, direction),
    showMpvOsd: (text: string) => deps.showMpvOsd(text),
    replayCurrentSubtitle: () => deps.replayCurrentSubtitle(),
    playNextSubtitle: () => deps.playNextSubtitle(),
    sendMpvCommand: (command: (string | number)[]) => deps.sendMpvCommand(command),
    isMpvConnected: () => deps.isMpvConnected(),
    hasRuntimeOptionsManager: () => deps.hasRuntimeOptionsManager(),
  });
}
