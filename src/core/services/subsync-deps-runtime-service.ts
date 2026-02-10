import {
  SubsyncManualPayload,
} from "../../types";
import {
  SubsyncRuntimeDeps,
} from "./subsync-runtime-service";
import { SubsyncResolvedConfig } from "../../subsync/utils";

interface MpvClientLike {
  connected: boolean;
  currentAudioStreamIndex: number | null;
  send: (payload: { command: (string | number)[] }) => void;
  requestProperty: (name: string) => Promise<unknown>;
}

export interface SubsyncDepsRuntimeOptions {
  getMpvClient: () => MpvClientLike | null;
  getResolvedSubsyncConfig: () => SubsyncResolvedConfig;
  isSubsyncInProgress: () => boolean;
  setSubsyncInProgress: (inProgress: boolean) => void;
  showMpvOsd: (text: string) => void;
  sendToVisibleOverlay: (
    channel: string,
    payload?: unknown,
    options?: { restoreOnModalClose?: "runtime-options" | "subsync" },
  ) => boolean;
}

export function createSubsyncRuntimeDepsService(
  options: SubsyncDepsRuntimeOptions,
): SubsyncRuntimeDeps {
  return {
    getMpvClient: options.getMpvClient,
    getResolvedSubsyncConfig: options.getResolvedSubsyncConfig,
    isSubsyncInProgress: options.isSubsyncInProgress,
    setSubsyncInProgress: options.setSubsyncInProgress,
    showMpvOsd: options.showMpvOsd,
    openManualPicker: (payload: SubsyncManualPayload) => {
      options.sendToVisibleOverlay("subsync:open-manual", payload, {
        restoreOnModalClose: "subsync",
      });
    },
  };
}
