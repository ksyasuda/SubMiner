import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import type { SessionNumericSelectionStartPayload } from '../../types/runtime';

type OverlayNumericSelectionWindow = {
  isDestroyed: () => boolean;
  isVisible: () => boolean;
  isFocused?: () => boolean;
  setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => void;
  focus: () => void;
  webContents: {
    isFocused?: () => boolean;
    focus: () => void;
    send: (channel: string, payload: SessionNumericSelectionStartPayload) => void;
  };
};

export function tryBeginVisibleOverlayNumericSelection(options: {
  actionId: SessionNumericSelectionStartPayload['actionId'];
  timeoutMs: number;
  getMainWindow: () => OverlayNumericSelectionWindow | null;
  getVisibleOverlayVisible: () => boolean;
}): boolean {
  if (!options.getVisibleOverlayVisible()) {
    return false;
  }

  const mainWindow = options.getMainWindow();
  if (!mainWindow || mainWindow.isDestroyed() || !mainWindow.isVisible()) {
    return false;
  }

  mainWindow.setIgnoreMouseEvents(false);
  if (typeof mainWindow.isFocused !== 'function' || !mainWindow.isFocused()) {
    mainWindow.focus();
  }
  if (
    typeof mainWindow.webContents.isFocused !== 'function' ||
    !mainWindow.webContents.isFocused()
  ) {
    mainWindow.webContents.focus();
  }
  mainWindow.webContents.send(IPC_CHANNELS.event.sessionNumericSelectionStart, {
    actionId: options.actionId,
    timeoutMs: options.timeoutMs,
  });
  return true;
}
