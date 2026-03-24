import type { YoutubePickerOpenPayload } from '../../types';
import type { OverlayHostedModal } from '../../shared/ipc/contracts';

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
  const sendPickerOpen = (): boolean =>
    deps.sendToActiveOverlayWindow('youtube:picker-open', payload, {
      restoreOnModalClose: YOUTUBE_PICKER_MODAL,
      preferModalWindow: true,
    });

  if (!sendPickerOpen()) {
    return false;
  }
  if (await deps.waitForModalOpen(YOUTUBE_PICKER_MODAL, YOUTUBE_PICKER_OPEN_TIMEOUT_MS)) {
    return true;
  }

  deps.logWarn(
    'YouTube subtitle picker did not acknowledge modal open on first attempt; retrying dedicated modal window.',
  );
  if (!sendPickerOpen()) {
    return false;
  }
  return await deps.waitForModalOpen(YOUTUBE_PICKER_MODAL, YOUTUBE_PICKER_OPEN_TIMEOUT_MS);
}
