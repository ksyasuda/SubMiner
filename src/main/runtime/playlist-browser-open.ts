import type { OverlayHostedModal } from '../../shared/ipc/contracts';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';

const PLAYLIST_BROWSER_MODAL: OverlayHostedModal = 'playlist-browser';

export function openPlaylistBrowser(deps: {
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
}): boolean {
  deps.ensureOverlayStartupPrereqs();
  deps.ensureOverlayWindowsReadyForVisibilityActions();
  return deps.sendToActiveOverlayWindow(IPC_CHANNELS.event.playlistBrowserOpen, undefined, {
    restoreOnModalClose: PLAYLIST_BROWSER_MODAL,
  });
}
