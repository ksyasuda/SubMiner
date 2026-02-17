import type { RuntimeOptionApplyResult, RuntimeOptionId } from "../types";
import { handleMpvCommandFromIpc } from "../core/services";
import { createMpvCommandRuntimeServiceDeps } from "./dependencies";
import { SPECIAL_COMMANDS } from "../config";

export interface MpvCommandFromIpcRuntimeDeps {
  triggerSubsyncFromConfig: () => void;
  openRuntimeOptionsPalette: () => void;
  cycleRuntimeOption: (
    id: RuntimeOptionId,
    direction: 1 | -1,
  ) => RuntimeOptionApplyResult;
  showMpvOsd: (text: string) => void;
  replayCurrentSubtitle: () => void;
  playNextSubtitle: () => void;
  sendMpvCommand: (command: (string | number)[]) => void;
  isMpvConnected: () => boolean;
  hasRuntimeOptionsManager: () => boolean;
}

export function handleMpvCommandFromIpcRuntime(
  command: (string | number)[],
  deps: MpvCommandFromIpcRuntimeDeps,
): void {
  handleMpvCommandFromIpc(
    command,
    createMpvCommandRuntimeServiceDeps({
      specialCommands: SPECIAL_COMMANDS,
      triggerSubsyncFromConfig: deps.triggerSubsyncFromConfig,
      openRuntimeOptionsPalette: deps.openRuntimeOptionsPalette,
      runtimeOptionsCycle: deps.cycleRuntimeOption,
      showMpvOsd: deps.showMpvOsd,
      mpvReplaySubtitle: deps.replayCurrentSubtitle,
      mpvPlayNextSubtitle: deps.playNextSubtitle,
      mpvSendCommand: deps.sendMpvCommand,
      isMpvConnected: deps.isMpvConnected,
      hasRuntimeOptionsManager: deps.hasRuntimeOptionsManager,
    }),
  );
}
