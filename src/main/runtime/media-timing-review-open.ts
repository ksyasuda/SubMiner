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
      // The review renderer regularly needs more than the 1.5 s the other modals allow; a
      // premature retry re-sends the payload and reloads the waveform for nothing.
      timeoutMs: 4_000,
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
