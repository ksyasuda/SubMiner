import type { BrowserWindow } from 'electron';

export type OverlayModalInputStateDeps = {
  getModalWindow: () => BrowserWindow | null;
  syncOverlayShortcutsForModal: (isActive: boolean) => void;
  syncOverlayVisibilityForModal: () => void;
};

export function createOverlayModalInputState(deps: OverlayModalInputStateDeps) {
  let modalInputExclusive = false;

  const handleModalInputStateChange = (isActive: boolean): void => {
    if (modalInputExclusive === isActive) {
      return;
    }

    modalInputExclusive = isActive;
    if (isActive) {
      const modalWindow = deps.getModalWindow();
      if (modalWindow && !modalWindow.isDestroyed()) {
        modalWindow.setIgnoreMouseEvents(false);
        modalWindow.setAlwaysOnTop(true, 'screen-saver', 1);
        modalWindow.focus();
        if (!modalWindow.webContents.isFocused()) {
          modalWindow.webContents.focus();
        }
      }
    }

    deps.syncOverlayShortcutsForModal(isActive);
    deps.syncOverlayVisibilityForModal();
  };

  return {
    getModalInputExclusive: (): boolean => modalInputExclusive,
    handleModalInputStateChange,
  };
}
