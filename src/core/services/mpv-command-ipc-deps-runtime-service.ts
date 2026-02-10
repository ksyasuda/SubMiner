import {
  RuntimeOptionApplyResult,
  RuntimeOptionId,
} from "../../types";
import {
  HandleMpvCommandFromIpcOptions,
} from "./ipc-command-service";
import { applyRuntimeOptionResultRuntimeService } from "./runtime-options-runtime-service";

interface RuntimeOptionsManagerLike {
  cycleOption: (
    id: RuntimeOptionId,
    direction: 1 | -1,
  ) => RuntimeOptionApplyResult;
}

export interface MpvCommandIpcDepsRuntimeOptions {
  specialCommands: HandleMpvCommandFromIpcOptions["specialCommands"];
  triggerSubsyncFromConfig: () => void;
  openRuntimeOptionsPalette: () => void;
  getRuntimeOptionsManager: () => RuntimeOptionsManagerLike | null;
  showMpvOsd: (text: string) => void;
  mpvReplaySubtitle: () => void;
  mpvPlayNextSubtitle: () => void;
  mpvSendCommand: (command: (string | number)[]) => void;
  isMpvConnected: () => boolean;
}

export function createMpvCommandIpcDepsRuntimeService(
  options: MpvCommandIpcDepsRuntimeOptions,
): HandleMpvCommandFromIpcOptions {
  return {
    specialCommands: options.specialCommands,
    triggerSubsyncFromConfig: options.triggerSubsyncFromConfig,
    openRuntimeOptionsPalette: options.openRuntimeOptionsPalette,
    runtimeOptionsCycle: (id, direction) => {
      const manager = options.getRuntimeOptionsManager();
      if (!manager) {
        return { ok: false, error: "Runtime options manager unavailable" };
      }
      return applyRuntimeOptionResultRuntimeService(
        manager.cycleOption(id, direction),
        options.showMpvOsd,
      );
    },
    showMpvOsd: options.showMpvOsd,
    mpvReplaySubtitle: options.mpvReplaySubtitle,
    mpvPlayNextSubtitle: options.mpvPlayNextSubtitle,
    mpvSendCommand: options.mpvSendCommand,
    isMpvConnected: options.isMpvConnected,
    hasRuntimeOptionsManager: () => options.getRuntimeOptionsManager() !== null,
  };
}
