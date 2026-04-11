import type { YoutubePickerOpenPayload } from '../../types';
import type { OverlayHostedModal } from '../../shared/ipc/contracts';
import { retryOverlayModalOpen } from './overlay-hosted-modal-open';

const YOUTUBE_PICKER_MODAL: OverlayHostedModal = 'youtube-track-picker';
const YOUTUBE_PICKER_OPEN_TIMEOUT_MS = 1500;

export async function openYoutubeTrackPicker(
  deps: {
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
  payload: YoutubePickerOpenPayload,
): Promise<boolean> {
  return await retryOverlayModalOpen(
    {
      waitForModalOpen: deps.waitForModalOpen,
      logWarn: deps.logWarn,
    },
    {
      modal: YOUTUBE_PICKER_MODAL,
      timeoutMs: YOUTUBE_PICKER_OPEN_TIMEOUT_MS,
      retryWarning:
        'YouTube subtitle picker did not acknowledge modal open on first attempt; retrying dedicated modal window.',
      sendOpen: () =>
        deps.sendToActiveOverlayWindow('youtube:picker-open', payload, {
          restoreOnModalClose: YOUTUBE_PICKER_MODAL,
          preferModalWindow: true,
        }),
    },
  );
}
