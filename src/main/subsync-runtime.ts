import { SubsyncResolvedConfig } from "../subsync/utils";
import type {
  SubsyncManualPayload,
  SubsyncManualRunRequest,
  SubsyncResult,
} from "../types";
import type { SubsyncRuntimeDeps } from "../core/services/subsync-runner";
import { createSubsyncRuntimeDeps } from "./dependencies";
import {
  runSubsyncManualFromIpcRuntime as runSubsyncManualFromIpcRuntimeCore,
  triggerSubsyncFromConfigRuntime as triggerSubsyncFromConfigRuntimeCore,
} from "../core/services";

export interface SubsyncRuntimeServiceInput {
  getMpvClient: SubsyncRuntimeDeps["getMpvClient"];
  getResolvedSubsyncConfig: () => SubsyncResolvedConfig;
  isSubsyncInProgress: SubsyncRuntimeDeps["isSubsyncInProgress"];
  setSubsyncInProgress: SubsyncRuntimeDeps["setSubsyncInProgress"];
  showMpvOsd: SubsyncRuntimeDeps["showMpvOsd"];
  openManualPicker: (payload: SubsyncManualPayload) => void;
}

export interface SubsyncRuntimeServiceStateInput {
  getMpvClient: SubsyncRuntimeServiceInput["getMpvClient"];
  getResolvedSubsyncConfig: SubsyncRuntimeServiceInput["getResolvedSubsyncConfig"];
  getSubsyncInProgress: () => boolean;
  setSubsyncInProgress: SubsyncRuntimeServiceInput["setSubsyncInProgress"];
  showMpvOsd: SubsyncRuntimeServiceInput["showMpvOsd"];
  openManualPicker: SubsyncRuntimeServiceInput["openManualPicker"];
}

export function createSubsyncRuntimeServiceInputFromState(
  input: SubsyncRuntimeServiceStateInput,
): SubsyncRuntimeServiceInput {
  return {
    getMpvClient: input.getMpvClient,
    getResolvedSubsyncConfig: input.getResolvedSubsyncConfig,
    isSubsyncInProgress: input.getSubsyncInProgress,
    setSubsyncInProgress: input.setSubsyncInProgress,
    showMpvOsd: input.showMpvOsd,
    openManualPicker: input.openManualPicker,
  };
}

export function createSubsyncRuntimeServiceDeps(
  params: SubsyncRuntimeServiceInput,
): SubsyncRuntimeDeps {
  return createSubsyncRuntimeDeps({
    getMpvClient: params.getMpvClient,
    getResolvedSubsyncConfig: params.getResolvedSubsyncConfig,
    isSubsyncInProgress: params.isSubsyncInProgress,
    setSubsyncInProgress: params.setSubsyncInProgress,
    showMpvOsd: params.showMpvOsd,
    openManualPicker: params.openManualPicker,
  });
}

export function triggerSubsyncFromConfigRuntime(
  params: SubsyncRuntimeServiceInput,
): Promise<void> {
  return triggerSubsyncFromConfigRuntimeCore(
    createSubsyncRuntimeServiceDeps(params),
  );
}

export async function runSubsyncManualFromIpcRuntime(
  request: SubsyncManualRunRequest,
  params: SubsyncRuntimeServiceInput,
): Promise<SubsyncResult> {
  return runSubsyncManualFromIpcRuntimeCore(
    request,
    createSubsyncRuntimeServiceDeps(params),
  );
}
