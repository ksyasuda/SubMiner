import type { OverlayHostedModal } from '../../shared/ipc/contracts';
import { openOverlayHostedModal, retryOverlayModalOpen } from './overlay-hosted-modal-open';

const RUNTIME_OPTIONS_MODAL: OverlayHostedModal = 'runtime-options';
const RUNTIME_OPTIONS_OPEN_TIMEOUT_MS = 1500;

export async function openRuntimeOptionsModal(deps: {
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
      modal: RUNTIME_OPTIONS_MODAL,
      timeoutMs: RUNTIME_OPTIONS_OPEN_TIMEOUT_MS,
      retryWarning:
        'Runtime options modal did not acknowledge modal open on first attempt; retrying dedicated modal window.',
      sendOpen: () =>
        openOverlayHostedModal(
          {
            ensureOverlayStartupPrereqs: deps.ensureOverlayStartupPrereqs,
            ensureOverlayWindowsReadyForVisibilityActions:
              deps.ensureOverlayWindowsReadyForVisibilityActions,
            sendToActiveOverlayWindow: deps.sendToActiveOverlayWindow,
          },
          {
            channel: 'runtime-options:open',
            modal: RUNTIME_OPTIONS_MODAL,
            preferModalWindow: true,
          },
        ),
    },
  );
}
