import {
  RuntimeOptionApplyResult,
  RuntimeOptionId,
  RuntimeOptionValue,
  SubsyncManualRunRequest,
  SubsyncResult,
} from "../../types";

export interface HandleMpvCommandFromIpcOptions {
  specialCommands: {
    SUBSYNC_TRIGGER: string;
    RUNTIME_OPTIONS_OPEN: string;
    RUNTIME_OPTION_CYCLE_PREFIX: string;
    REPLAY_SUBTITLE: string;
    PLAY_NEXT_SUBTITLE: string;
  };
  triggerSubsyncFromConfig: () => void;
  openRuntimeOptionsPalette: () => void;
  runtimeOptionsCycle: (
    id: RuntimeOptionId,
    direction: 1 | -1,
  ) => RuntimeOptionApplyResult;
  showMpvOsd: (text: string) => void;
  mpvReplaySubtitle: () => void;
  mpvPlayNextSubtitle: () => void;
  mpvSendCommand: (command: (string | number)[]) => void;
  isMpvConnected: () => boolean;
  hasRuntimeOptionsManager: () => boolean;
}

export function handleMpvCommandFromIpcService(
  command: (string | number)[],
  options: HandleMpvCommandFromIpcOptions,
): void {
  const first = typeof command[0] === "string" ? command[0] : "";
  if (first === options.specialCommands.SUBSYNC_TRIGGER) {
    options.triggerSubsyncFromConfig();
    return;
  }

  if (first === options.specialCommands.RUNTIME_OPTIONS_OPEN) {
    options.openRuntimeOptionsPalette();
    return;
  }

  if (first.startsWith(options.specialCommands.RUNTIME_OPTION_CYCLE_PREFIX)) {
    if (!options.hasRuntimeOptionsManager()) return;
    const [, idToken, directionToken] = first.split(":");
    const id = idToken as RuntimeOptionId;
    const direction: 1 | -1 = directionToken === "prev" ? -1 : 1;
    const result = options.runtimeOptionsCycle(id, direction);
    if (!result.ok && result.error) {
      options.showMpvOsd(result.error);
    }
    return;
  }

  if (options.isMpvConnected()) {
    if (first === options.specialCommands.REPLAY_SUBTITLE) {
      options.mpvReplaySubtitle();
    } else if (first === options.specialCommands.PLAY_NEXT_SUBTITLE) {
      options.mpvPlayNextSubtitle();
    } else {
      options.mpvSendCommand(command);
    }
  }
}

export async function runSubsyncManualFromIpcService(
  request: SubsyncManualRunRequest,
  options: {
    isSubsyncInProgress: () => boolean;
    setSubsyncInProgress: (inProgress: boolean) => void;
    showMpvOsd: (text: string) => void;
    runWithSpinner: (task: () => Promise<SubsyncResult>) => Promise<SubsyncResult>;
    runSubsyncManual: (request: SubsyncManualRunRequest) => Promise<SubsyncResult>;
  },
): Promise<SubsyncResult> {
  if (options.isSubsyncInProgress()) {
    const busy = "Subsync already running";
    options.showMpvOsd(busy);
    return { ok: false, message: busy };
  }
  try {
    options.setSubsyncInProgress(true);
    const result = await options.runWithSpinner(() =>
      options.runSubsyncManual(request),
    );
    options.showMpvOsd(result.message);
    return result;
  } catch (error) {
    const message = `Subsync failed: ${(error as Error).message}`;
    options.showMpvOsd(message);
    return { ok: false, message };
  } finally {
    options.setSubsyncInProgress(false);
  }
}
