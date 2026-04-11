import type { OverlayHostedModal } from '../../shared/ipc/contracts';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { openOverlayHostedModal, retryOverlayModalOpen } from './overlay-hosted-modal-open';

const CONTROLLER_SELECT_MODAL: OverlayHostedModal = 'controller-select';
const CONTROLLER_SELECT_OPEN_TIMEOUT_MS = 1500;

export async function openControllerSelectModal(deps: {
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
      modal: CONTROLLER_SELECT_MODAL,
      timeoutMs: CONTROLLER_SELECT_OPEN_TIMEOUT_MS,
      retryWarning:
        'Controller select modal did not acknowledge modal open on first attempt; retrying dedicated modal window.',
      sendOpen: () =>
        openOverlayHostedModal(
          {
            ensureOverlayStartupPrereqs: deps.ensureOverlayStartupPrereqs,
            ensureOverlayWindowsReadyForVisibilityActions:
              deps.ensureOverlayWindowsReadyForVisibilityActions,
            sendToActiveOverlayWindow: deps.sendToActiveOverlayWindow,
          },
          {
            channel: IPC_CHANNELS.event.controllerSelectOpen,
            modal: CONTROLLER_SELECT_MODAL,
            preferModalWindow: true,
          },
        ),
    },
  );
}
