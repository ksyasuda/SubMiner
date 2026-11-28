import * as path from 'node:path';
import type { MpvIpcClient } from '../../core/services/mpv';
import {
  openSubsyncManualPickerRuntime,
  runSubsyncManualFromIpcRuntime,
  triggerSubsyncFromConfigRuntime,
} from '../../core/services/subsync-runner';
import { expandSubtitleGenerationPath } from '../../core/services/subtitle-generation-files';
import { SUBTITLE_GENERATION_VAD_MODEL } from '../../shared/subtitle-generation-vad-model';
import type {
  SubsyncResult,
  SubsyncManualPayload,
  SubsyncManualRunRequest,
  ResolvedConfig,
} from '../../types';
import { fileExists, findSubsyncExecutable, getSubsyncConfig } from '../../subsync/utils';
import { createSubsyncRuntimeServiceInputFromState } from '../subsync-runtime';

export type MainSubsyncRuntimeDeps = {
  getMpvClient: () => MpvIpcClient | null;
  getResolvedConfig: () => ResolvedConfig;
  getSubtitleGenerationModelDirectory?: () => string;
  canPickReferenceFile?: () => boolean;
  getSubsyncInProgress: () => boolean;
  setSubsyncInProgress: (inProgress: boolean) => void;
  showMpvOsd: (text: string) => void;
  openManualPicker: (payload: SubsyncManualPayload) => void;
};

export function createMainSubsyncRuntime(deps: MainSubsyncRuntimeDeps): {
  triggerFromConfig: () => Promise<void>;
  openManualPicker: () => Promise<void>;
  runManualFromIpc: (request: SubsyncManualRunRequest) => Promise<SubsyncResult>;
} {
  const resolveVadPaths = () => {
    const config = deps.getResolvedConfig();
    const gen = config.subtitleGeneration;
    const vadExecutablePath = findSubsyncExecutable(gen?.vadPath, [
      'whisper-vad-speech-segments',
      'vad-speech-segments',
    ]);
    let vadModelPath: string | null = null;
    const configuredModel = gen?.vadModelPath?.trim() ?? '';
    if (configuredModel) {
      const expanded = path.resolve(expandSubtitleGenerationPath(configuredModel));
      if (fileExists(expanded)) vadModelPath = expanded;
    } else if (deps.getSubtitleGenerationModelDirectory) {
      const managedModel = path.join(
        deps.getSubtitleGenerationModelDirectory(),
        SUBTITLE_GENERATION_VAD_MODEL.filename,
      );
      if (fileExists(managedModel)) vadModelPath = managedModel;
    }
    return { vadExecutablePath, vadModelPath };
  };

  const getRuntimeServiceParams = () =>
    createSubsyncRuntimeServiceInputFromState({
      getMpvClient: () => deps.getMpvClient(),
      getResolvedSubsyncConfig: () => getSubsyncConfig(deps.getResolvedConfig().subsync),
      getVadPaths: resolveVadPaths,
      canPickReferenceFile: deps.canPickReferenceFile,
      getSubsyncInProgress: () => deps.getSubsyncInProgress(),
      setSubsyncInProgress: (inProgress: boolean) => deps.setSubsyncInProgress(inProgress),
      showMpvOsd: (text: string) => deps.showMpvOsd(text),
      openManualPicker: (payload: SubsyncManualPayload) => deps.openManualPicker(payload),
    });

  return {
    triggerFromConfig: async (): Promise<void> => {
      await triggerSubsyncFromConfigRuntime(getRuntimeServiceParams());
    },
    openManualPicker: async (): Promise<void> => {
      await openSubsyncManualPickerRuntime(getRuntimeServiceParams());
    },
    runManualFromIpc: async (request: SubsyncManualRunRequest): Promise<SubsyncResult> =>
      runSubsyncManualFromIpcRuntime(request, getRuntimeServiceParams()),
  };
}
