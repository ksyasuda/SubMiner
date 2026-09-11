import type { IpcMain, WebContents } from 'electron';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { isSubtitleGenerationModelId } from '../../shared/subtitle-generation-model-catalog';
import type { createSubtitleGenerationRuntime } from './subtitle-generation-runtime';

export function registerSubtitleGenerationIpc(deps: {
  ipc: Pick<IpcMain, 'handle'>;
  isAllowedSender: (sender: WebContents) => boolean;
  openModal: () => Promise<boolean>;
  runtime: ReturnType<typeof createSubtitleGenerationRuntime>;
}): void {
  const handlers = [
    [IPC_CHANNELS.request.requestSubtitleGenerationOpen, () => deps.openModal()],
    [IPC_CHANNELS.request.getSubtitleGenerationStatus, () => deps.runtime.getStatus()],
    [
      IPC_CHANNELS.request.selectSubtitleGenerationModel,
      (model: unknown) => {
        if (!isSubtitleGenerationModelId(model))
          throw new Error('Unknown subtitle generation model.');
        return deps.runtime.selectModel(model);
      },
    ],
    [IPC_CHANNELS.request.startSubtitleGeneration, () => deps.runtime.start()],
    [IPC_CHANNELS.request.downloadSubtitleGenerationModel, () => deps.runtime.download()],
    [IPC_CHANNELS.request.downloadSubtitleGenerationVadModel, () => deps.runtime.downloadVad()],
    [
      IPC_CHANNELS.request.setSubtitleGenerationVadEnabled,
      (enabled: unknown) => {
        if (typeof enabled !== 'boolean')
          throw new Error('Speech detection selection must be a boolean.');
        return deps.runtime.setVadEnabled(enabled);
      },
    ],
    [IPC_CHANNELS.request.cancelSubtitleGeneration, () => deps.runtime.cancel()],
  ] as const;
  for (const [channel, handler] of handlers) {
    deps.ipc.handle(channel, (event, payload: unknown) => {
      if (!deps.isAllowedSender(event.sender))
        throw new Error('Subtitle generation is only available from the overlay.');
      return handler(payload);
    });
  }
}
