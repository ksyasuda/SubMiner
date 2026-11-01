import type { OverlayHostedModal } from '../../shared/ipc/contracts';
import type { OverlayModalRuntime } from '../overlay-runtime';

export function openOverlayHostedModal(
  deps: {
    ensureOverlayStartupPrereqs: () => void;
    ensureOverlayWindowsReadyForVisibilityActions: () => void;
    sendToActiveOverlayWindow: OverlayModalRuntime['sendToActiveOverlayWindow'];
  },
  input: {
    channel: string;
    modal: OverlayHostedModal;
    payload?: unknown;
    preferModalWindow?: boolean;
    signal?: AbortSignal;
  },
): boolean {
  deps.ensureOverlayStartupPrereqs();
  deps.ensureOverlayWindowsReadyForVisibilityActions();
  return deps.sendToActiveOverlayWindow(input.channel, input.payload, {
    restoreOnModalClose: input.modal,
    preferModalWindow: input.preferModalWindow,
    ...(input.signal ? { signal: input.signal } : {}),
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
    signal?: AbortSignal;
  },
): Promise<boolean> {
  if (input.signal?.aborted || !input.sendOpen()) {
    return false;
  }

  const opened = await deps.waitForModalOpen(input.modal, input.timeoutMs);
  if (input.signal?.aborted) return false;
  if (opened) {
    return true;
  }

  deps.logWarn(input.retryWarning);
  if (input.signal?.aborted || !input.sendOpen()) {
    return false;
  }

  const retryOpened = await deps.waitForModalOpen(input.modal, input.timeoutMs);
  return !input.signal?.aborted && retryOpened;
}
