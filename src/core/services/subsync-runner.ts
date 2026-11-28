import { SubsyncManualPayload, SubsyncManualRunRequest, SubsyncResult } from '../../types';
import { SubsyncResolvedConfig } from '../../subsync/utils';
import { runSubsyncManualFromIpc } from './ipc-command';
import {
  SubsyncCancelledError,
  SubsyncVadPaths,
  TriggerSubsyncFromConfigDeps,
  openSubsyncManualPicker,
  runSubsyncManual,
  runWithMediaChangeAbort,
  triggerSubsyncFromConfig,
} from './subsync';

const AUTOSUBSYNC_SPINNER_FRAMES = ['|', '/', '-', '\\'];

interface MpvClientLike {
  connected: boolean;
  currentAudioStreamIndex: number | null;
  send: (payload: { command: (string | number)[] }) => void;
  requestProperty: (name: string) => Promise<unknown>;
}

export interface SubsyncRuntimeDeps {
  getMpvClient: () => MpvClientLike | null;
  getResolvedSubsyncConfig: () => SubsyncResolvedConfig;
  getVadPaths?: () => SubsyncVadPaths;
  canPickReferenceFile?: () => boolean;
  isSubsyncInProgress: () => boolean;
  setSubsyncInProgress: (inProgress: boolean) => void;
  showMpvOsd: (text: string) => void;
  openManualPicker: (payload: SubsyncManualPayload) => void;
}

async function runWithSubsyncSpinnerService<T>(
  task: () => Promise<T>,
  showMpvOsd: (text: string) => void,
  label = 'Subsync: syncing',
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

function buildTriggerSubsyncDeps(deps: SubsyncRuntimeDeps): TriggerSubsyncFromConfigDeps {
  return {
    getMpvClient: deps.getMpvClient,
    getResolvedConfig: deps.getResolvedSubsyncConfig,
    getVadPaths: deps.getVadPaths,
    canPickReferenceFile: deps.canPickReferenceFile,
    isSubsyncInProgress: deps.isSubsyncInProgress,
    setSubsyncInProgress: deps.setSubsyncInProgress,
    showMpvOsd: deps.showMpvOsd,
    runWithSubsyncSpinner: <T>(task: () => Promise<T>) =>
      runWithSubsyncSpinnerService(task, deps.showMpvOsd),
    openManualPicker: deps.openManualPicker,
  };
}

export async function triggerSubsyncFromConfigRuntime(deps: SubsyncRuntimeDeps): Promise<void> {
  await triggerSubsyncFromConfig(buildTriggerSubsyncDeps(deps));
}

export async function openSubsyncManualPickerRuntime(deps: SubsyncRuntimeDeps): Promise<void> {
  const triggerDeps = buildTriggerSubsyncDeps(deps);
  if (triggerDeps.isSubsyncInProgress()) {
    triggerDeps.showMpvOsd('Subsync already running');
    return;
  }
  triggerDeps.setSubsyncInProgress(true);
  try {
    const opened = await triggerDeps.runWithSubsyncSpinner(() =>
      runWithMediaChangeAbort(triggerDeps, (runDeps) =>
        openSubsyncManualPicker(runDeps, {
          guidanceMessage: null,
        }),
      ),
    );
    if (opened) {
      triggerDeps.showMpvOsd('Subsync: choose engine and subtitles');
    }
  } catch (error) {
    triggerDeps.showMpvOsd(
      error instanceof SubsyncCancelledError
        ? error.message
        : `Subsync failed: ${(error as Error).message}`,
    );
  } finally {
    triggerDeps.setSubsyncInProgress(false);
  }
}

export async function runSubsyncManualFromIpcRuntime(
  request: SubsyncManualRunRequest,
  deps: SubsyncRuntimeDeps,
): Promise<SubsyncResult> {
  const triggerDeps = buildTriggerSubsyncDeps(deps);
  return runSubsyncManualFromIpc(request, {
    isSubsyncInProgress: triggerDeps.isSubsyncInProgress,
    setSubsyncInProgress: triggerDeps.setSubsyncInProgress,
    showMpvOsd: triggerDeps.showMpvOsd,
    runWithSpinner: (task) => triggerDeps.runWithSubsyncSpinner(() => task()),
    runSubsyncManual: (subsyncRequest) => runSubsyncManual(subsyncRequest, triggerDeps),
  });
}
