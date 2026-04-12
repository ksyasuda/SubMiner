import type { OverlayHostedModal } from '../../shared/ipc/contracts';

export function openOverlayHostedModal(
  deps: {
    ensureOverlayStartupPrereqs: () => void;
    ensureOverlayWindowsReadyForVisibilityActions: () => void;
    sendToActiveOverlayWindow: (
      channel: string,
      payload?: unknown,
      runtimeOptions?: {
        restoreOnModalClose?: OverlayHostedModal;
        preferModalWindow?: boolean;
      },
    ) => boolean;
  },
  input: {
    channel: string;
    modal: OverlayHostedModal;
    payload?: unknown;
    preferModalWindow?: boolean;
  },
): boolean {
  deps.ensureOverlayStartupPrereqs();
  deps.ensureOverlayWindowsReadyForVisibilityActions();
  return deps.sendToActiveOverlayWindow(input.channel, input.payload, {
    restoreOnModalClose: input.modal,
    preferModalWindow: input.preferModalWindow,
  });
}

export async function retryOverlayModalOpen(
  deps: {
    waitForModalOpen: (modal: OverlayHostedModal, timeoutMs: number) => Promise<boolean>;
    logWarn: (message: string) => void;
  },
  input: {
    modal: OverlayHostedModal;
    timeoutMs: number;
    retryWarning: string;
    sendOpen: () => boolean;
  },
): Promise<boolean> {
  if (!input.sendOpen()) {
    return false;
  }

  if (await deps.waitForModalOpen(input.modal, input.timeoutMs)) {
    return true;
  }

  deps.logWarn(input.retryWarning);
  if (!input.sendOpen()) {
    return false;
  }

  return await deps.waitForModalOpen(input.modal, input.timeoutMs);
}
