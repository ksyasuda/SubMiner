import type { BrowserWindow } from 'electron';

function requestOverlayApplicationFocus(): void {
  try {
    const electron = require('electron') as {
      app?: {
        focus?: (options?: { steal?: boolean }) => void;
      };
    };
    electron.app?.focus?.({ steal: true });
  } catch {
    // Ignore focus-steal failures in non-Electron test environments.
  }
}

function setWindowFocusable(window: BrowserWindow): void {
  const maybeFocusableWindow = window as BrowserWindow & {
    setFocusable?: (focusable: boolean) => void;
  };
  maybeFocusableWindow.setFocusable?.(true);
}

export type OverlayModalInputStateDeps = {
  getModalWindow: () => BrowserWindow | null;
  syncOverlayShortcutsForModal: (isActive: boolean) => void;
  syncOverlayVisibilityForModal: () => void;
  restoreMainWindowFocus?: () => void;
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
        setWindowFocusable(modalWindow);
        requestOverlayApplicationFocus();
        modalWindow.setIgnoreMouseEvents(false);
        modalWindow.setAlwaysOnTop(true, 'screen-saver', 3);
        modalWindow.focus();
        if (!modalWindow.webContents.isFocused()) {
          modalWindow.webContents.focus();
        }
      }
    }

    deps.syncOverlayShortcutsForModal(isActive);
    deps.syncOverlayVisibilityForModal();
    if (!isActive) {
      deps.restoreMainWindowFocus?.();
    }
  };

  return {
    getModalInputExclusive: (): boolean => modalInputExclusive,
    handleModalInputStateChange,
  };
}
