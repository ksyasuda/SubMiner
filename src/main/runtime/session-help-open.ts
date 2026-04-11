import type { OverlayHostedModal } from '../../shared/ipc/contracts';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { openOverlayHostedModal, retryOverlayModalOpen } from './overlay-hosted-modal-open';

const SESSION_HELP_MODAL: OverlayHostedModal = 'session-help';
const SESSION_HELP_OPEN_TIMEOUT_MS = 1500;

export async function openSessionHelpModal(deps: {
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
  waitForModalOpen: (modal: OverlayHostedModal, timeoutMs: number) => Promise<boolean>;
  logWarn: (message: string) => void;
}): Promise<boolean> {
  return await retryOverlayModalOpen(
    {
      waitForModalOpen: deps.waitForModalOpen,
      logWarn: deps.logWarn,
    },
    {
      modal: SESSION_HELP_MODAL,
      timeoutMs: SESSION_HELP_OPEN_TIMEOUT_MS,
      retryWarning:
        'Session help modal did not acknowledge modal open on first attempt; retrying dedicated modal window.',
      sendOpen: () =>
        openOverlayHostedModal(
          {
            ensureOverlayStartupPrereqs: deps.ensureOverlayStartupPrereqs,
            ensureOverlayWindowsReadyForVisibilityActions:
              deps.ensureOverlayWindowsReadyForVisibilityActions,
            sendToActiveOverlayWindow: deps.sendToActiveOverlayWindow,
          },
          {
            channel: IPC_CHANNELS.event.sessionHelpOpen,
            modal: SESSION_HELP_MODAL,
            preferModalWindow: true,
          },
        ),
    },
  );
}
