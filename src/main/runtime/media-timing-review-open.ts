import { IPC_CHANNELS, type OverlayHostedModal } from '../../shared/ipc/contracts';
import type { MediaTimingReviewOpenPayload } from '../../types/anki';
import { openOverlayHostedModal, retryOverlayModalOpen } from './overlay-hosted-modal-open';

const MODAL: OverlayHostedModal = 'media-timing-review';

export async function openMediaTimingReviewModal(
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
  payload: MediaTimingReviewOpenPayload,
): Promise<boolean> {
  return await retryOverlayModalOpen(
    { waitForModalOpen: deps.waitForModalOpen, logWarn: deps.logWarn },
    {
      modal: MODAL,
      timeoutMs: 1_500,
      retryWarning:
        'Media timing review did not acknowledge modal open; retrying the dedicated modal window.',
      sendOpen: () =>
        openOverlayHostedModal(deps, {
          channel: IPC_CHANNELS.event.mediaTimingReviewOpen,
          modal: MODAL,
          payload,
          preferModalWindow: true,
        }),
    },
  );
}
