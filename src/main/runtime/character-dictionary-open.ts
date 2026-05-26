import type { OverlayHostedModal } from '../../shared/ipc/contracts';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { openOverlayHostedModal, retryOverlayModalOpen } from './overlay-hosted-modal-open';

const CHARACTER_DICTIONARY_MODAL: OverlayHostedModal = 'character-dictionary';
const CHARACTER_DICTIONARY_OPEN_TIMEOUT_MS = 1500;

async function openCharacterDictionaryModalChannel(deps: {
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
  channel: string;
  retryWarning: string;
}): Promise<boolean> {
  return await retryOverlayModalOpen(
    {
      waitForModalOpen: deps.waitForModalOpen,
      logWarn: deps.logWarn,
    },
    {
      modal: CHARACTER_DICTIONARY_MODAL,
      timeoutMs: CHARACTER_DICTIONARY_OPEN_TIMEOUT_MS,
      retryWarning: deps.retryWarning,
      sendOpen: () =>
        openOverlayHostedModal(
          {
            ensureOverlayStartupPrereqs: deps.ensureOverlayStartupPrereqs,
            ensureOverlayWindowsReadyForVisibilityActions:
              deps.ensureOverlayWindowsReadyForVisibilityActions,
            sendToActiveOverlayWindow: deps.sendToActiveOverlayWindow,
          },
          {
            channel: deps.channel,
            modal: CHARACTER_DICTIONARY_MODAL,
            preferModalWindow: true,
          },
        ),
    },
  );
}

type OpenCharacterDictionaryModalDeps = Omit<
  Parameters<typeof openCharacterDictionaryModalChannel>[0],
  'channel' | 'retryWarning'
>;

export async function openCharacterDictionaryManagerModal(
  deps: OpenCharacterDictionaryModalDeps,
): Promise<boolean> {
  return await openCharacterDictionaryModalChannel({
    ...deps,
    channel: IPC_CHANNELS.event.characterDictionaryManagerOpen,
    retryWarning:
      'Character dictionary manager did not acknowledge modal open on first attempt; retrying dedicated modal window.',
  });
}
