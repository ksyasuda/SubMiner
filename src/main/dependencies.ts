import {
  RuntimeOptionId,
  RuntimeOptionValue,
  SubsyncManualPayload,
} from "../types";
import { SubsyncResolvedConfig } from "../subsync/utils";
import type { SubsyncRuntimeDeps } from "../core/services/subsync-runner-service";
import {
  cycleRuntimeOptionFromIpcRuntimeService,
  setRuntimeOptionFromIpcRuntimeService,
} from "../core/services/runtime-options-ipc-service";
import { RuntimeOptionsManager } from "../runtime-options";

export interface RuntimeOptionsIpcDepsParams {
  getRuntimeOptionsManager: () => RuntimeOptionsManager | null;
  showMpvOsd: (text: string) => void;
}

export interface SubsyncRuntimeDepsParams {
  getMpvClient: () => ReturnType<SubsyncRuntimeDeps["getMpvClient"]>;
  getResolvedSubsyncConfig: () => SubsyncResolvedConfig;
  isSubsyncInProgress: () => boolean;
  setSubsyncInProgress: (inProgress: boolean) => void;
  showMpvOsd: (text: string) => void;
  openManualPicker: (payload: SubsyncManualPayload) => void;
}

export function createRuntimeOptionsIpcDeps(params: RuntimeOptionsIpcDepsParams): {
  setRuntimeOption: (id: string, value: unknown) => unknown;
  cycleRuntimeOption: (id: string, direction: 1 | -1) => unknown;
} {
  return {
    setRuntimeOption: (id, value) =>
      setRuntimeOptionFromIpcRuntimeService(
        params.getRuntimeOptionsManager(),
        id as RuntimeOptionId,
        value as RuntimeOptionValue,
        (text) => params.showMpvOsd(text),
      ),
    cycleRuntimeOption: (id, direction) =>
      cycleRuntimeOptionFromIpcRuntimeService(
        params.getRuntimeOptionsManager(),
        id as RuntimeOptionId,
        direction,
        (text) => params.showMpvOsd(text),
      ),
  };
}

export function createSubsyncRuntimeDeps(params: SubsyncRuntimeDepsParams): SubsyncRuntimeDeps {
  return {
    getMpvClient: params.getMpvClient,
    getResolvedSubsyncConfig: params.getResolvedSubsyncConfig,
    isSubsyncInProgress: params.isSubsyncInProgress,
    setSubsyncInProgress: params.setSubsyncInProgress,
    showMpvOsd: params.showMpvOsd,
    openManualPicker: params.openManualPicker,
  };
}
