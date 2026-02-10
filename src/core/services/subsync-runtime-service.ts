import {
  SubsyncManualPayload,
  SubsyncManualRunRequest,
  SubsyncResult,
} from "../../types";
import { SubsyncResolvedConfig } from "../../subsync/utils";
import { runSubsyncManualFromIpcService } from "./ipc-command-service";
import {
  TriggerSubsyncFromConfigDeps,
  runSubsyncManualService,
  triggerSubsyncFromConfigService,
} from "./subsync-service";

const AUTOSUBSYNC_SPINNER_FRAMES = ["|", "/", "-", "\\"];

interface MpvClientLike {
  connected: boolean;
  currentAudioStreamIndex: number | null;
  send: (payload: { command: (string | number)[] }) => void;
  requestProperty: (name: string) => Promise<unknown>;
}

export interface SubsyncRuntimeDeps {
  getMpvClient: () => MpvClientLike | null;
  getResolvedSubsyncConfig: () => SubsyncResolvedConfig;
  isSubsyncInProgress: () => boolean;
  setSubsyncInProgress: (inProgress: boolean) => void;
  showMpvOsd: (text: string) => void;
  openManualPicker: (payload: SubsyncManualPayload) => void;
}

async function runWithSubsyncSpinnerService<T>(
  task: () => Promise<T>,
  showMpvOsd: (text: string) => void,
  label = "Subsync: syncing",
): Promise<T> {
  let frame = 0;
  showMpvOsd(`${label} ${AUTOSUBSYNC_SPINNER_FRAMES[0]}`);
  const timer = setInterval(() => {
    frame = (frame + 1) % AUTOSUBSYNC_SPINNER_FRAMES.length;
    showMpvOsd(`${label} ${AUTOSUBSYNC_SPINNER_FRAMES[frame]}`);
  }, 150);
  try {
    return await task();
  } finally {
    clearInterval(timer);
  }
}

function buildTriggerSubsyncDeps(
  deps: SubsyncRuntimeDeps,
): TriggerSubsyncFromConfigDeps {
  return {
    getMpvClient: deps.getMpvClient,
    getResolvedConfig: deps.getResolvedSubsyncConfig,
    isSubsyncInProgress: deps.isSubsyncInProgress,
    setSubsyncInProgress: deps.setSubsyncInProgress,
    showMpvOsd: deps.showMpvOsd,
    runWithSubsyncSpinner: <T>(task: () => Promise<T>) =>
      runWithSubsyncSpinnerService(task, deps.showMpvOsd),
    openManualPicker: deps.openManualPicker,
  };
}

export async function triggerSubsyncFromConfigRuntimeService(
  deps: SubsyncRuntimeDeps,
): Promise<void> {
  await triggerSubsyncFromConfigService(buildTriggerSubsyncDeps(deps));
}

export async function runSubsyncManualFromIpcRuntimeService(
  request: SubsyncManualRunRequest,
  deps: SubsyncRuntimeDeps,
): Promise<SubsyncResult> {
  const triggerDeps = buildTriggerSubsyncDeps(deps);
  return runSubsyncManualFromIpcService(request, {
    isSubsyncInProgress: triggerDeps.isSubsyncInProgress,
    setSubsyncInProgress: triggerDeps.setSubsyncInProgress,
    showMpvOsd: triggerDeps.showMpvOsd,
    runWithSpinner: (task) =>
      triggerDeps.runWithSubsyncSpinner(() => task()),
    runSubsyncManual: (subsyncRequest) =>
      runSubsyncManualService(subsyncRequest, triggerDeps),
  });
}
