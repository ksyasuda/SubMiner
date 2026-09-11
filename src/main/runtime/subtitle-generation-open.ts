import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { openOverlayHostedModal, retryOverlayModalOpen } from './overlay-hosted-modal-open';

export function openSubtitleGenerationModal(
  deps: Parameters<typeof openOverlayHostedModal>[0] & Parameters<typeof retryOverlayModalOpen>[0],
): Promise<boolean> {
  return retryOverlayModalOpen(deps, {
    modal: 'subtitle-generation',
    timeoutMs: 1500,
    retryWarning: 'Subtitle generation modal did not acknowledge opening; retrying.',
    sendOpen: () =>
      openOverlayHostedModal(deps, {
        channel: IPC_CHANNELS.event.subtitleGenerationOpen,
        modal: 'subtitle-generation',
        preferModalWindow: true,
      }),
  });
}
