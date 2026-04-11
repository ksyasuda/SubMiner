import type { OverlayHostedModal } from '../../shared/ipc/contracts';

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
  const sendOpen = (): boolean => {
    deps.ensureOverlayStartupPrereqs();
    deps.ensureOverlayWindowsReadyForVisibilityActions();
    return deps.sendToActiveOverlayWindow('runtime-options:open', undefined, {
      restoreOnModalClose: RUNTIME_OPTIONS_MODAL,
      preferModalWindow: true,
    });
  };

  if (!sendOpen()) {
    return false;
  }

  if (await deps.waitForModalOpen(RUNTIME_OPTIONS_MODAL, RUNTIME_OPTIONS_OPEN_TIMEOUT_MS)) {
    return true;
  }

  deps.logWarn(
    'Runtime options modal did not acknowledge modal open on first attempt; retrying dedicated modal window.',
  );
  if (!sendOpen()) {
    return false;
  }

  return await deps.waitForModalOpen(RUNTIME_OPTIONS_MODAL, RUNTIME_OPTIONS_OPEN_TIMEOUT_MS);
}
