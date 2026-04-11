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
