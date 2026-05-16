import type { OverlayHostedModal } from '../../shared/ipc/contracts';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import type { SubsyncManualPayload } from '../../types';
import { openOverlayHostedModal, retryOverlayModalOpen } from './overlay-hosted-modal-open';

const SUBSYNC_MODAL: OverlayHostedModal = 'subsync';
const SUBSYNC_OPEN_TIMEOUT_MS = 1500;

export async function openSubsyncManualModal(
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
    waitForModalOpen: (modal: OverlayHostedModal, timeoutMs: number) => Promise<boolean>;
    logWarn: (message: string) => void;
  },
  payload: SubsyncManualPayload,
): Promise<boolean> {
  return await retryOverlayModalOpen(
    {
      waitForModalOpen: deps.waitForModalOpen,
      logWarn: deps.logWarn,
    },
    {
      modal: SUBSYNC_MODAL,
      timeoutMs: SUBSYNC_OPEN_TIMEOUT_MS,
      retryWarning:
        'Subsync modal did not acknowledge modal open on first attempt; retrying dedicated modal window.',
      sendOpen: () =>
        openOverlayHostedModal(
          {
            ensureOverlayStartupPrereqs: deps.ensureOverlayStartupPrereqs,
            ensureOverlayWindowsReadyForVisibilityActions:
              deps.ensureOverlayWindowsReadyForVisibilityActions,
            sendToActiveOverlayWindow: deps.sendToActiveOverlayWindow,
          },
          {
            channel: IPC_CHANNELS.event.subsyncOpenManual,
            modal: SUBSYNC_MODAL,
            payload,
            preferModalWindow: true,
          },
        ),
    },
  );
}
