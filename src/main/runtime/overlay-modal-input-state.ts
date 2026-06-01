import type { BrowserWindow } from 'electron';

type VisibilitySyncTimeout = NonNullable<Parameters<typeof globalThis.clearTimeout>[0]>;
const POST_RESTORE_VISIBILITY_SYNC_DELAYS_MS = [50, 150, 300, 600, 1000] as const;

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
  schedulePostRestoreVisibilitySync?: (
    callback: () => void,
    delayMs: number,
  ) => VisibilitySyncTimeout;
  clearPostRestoreVisibilitySync?: (timeout: VisibilitySyncTimeout) => void;
};

export function createOverlayModalInputState(deps: OverlayModalInputStateDeps) {
  let modalInputExclusive = false;
  let postRestoreVisibilitySyncTimeouts: VisibilitySyncTimeout[] = [];
  const schedulePostRestoreVisibilitySync =
    deps.schedulePostRestoreVisibilitySync ?? globalThis.setTimeout;
  const clearPostRestoreVisibilitySync =
    deps.clearPostRestoreVisibilitySync ?? globalThis.clearTimeout;

  const clearPostRestoreVisibilitySyncBurst = (): void => {
    for (const timeout of postRestoreVisibilitySyncTimeouts) {
      clearPostRestoreVisibilitySync(timeout);
    }
    postRestoreVisibilitySyncTimeouts = [];
  };

  const schedulePostRestoreVisibilitySyncBurst = (): void => {
    clearPostRestoreVisibilitySyncBurst();
    for (const delayMs of POST_RESTORE_VISIBILITY_SYNC_DELAYS_MS) {
      const timeout = schedulePostRestoreVisibilitySync(() => {
        postRestoreVisibilitySyncTimeouts = postRestoreVisibilitySyncTimeouts.filter(
          (candidate) => candidate !== timeout,
        );
        deps.syncOverlayVisibilityForModal();
      }, delayMs);
      (timeout as { unref?: () => void }).unref?.();
      postRestoreVisibilitySyncTimeouts.push(timeout);
    }
  };

  const handleModalInputStateChange = (isActive: boolean): void => {
    if (modalInputExclusive === isActive) {
      return;
    }

    clearPostRestoreVisibilitySyncBurst();
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
      if (deps.restoreMainWindowFocus) {
        deps.syncOverlayVisibilityForModal();
        schedulePostRestoreVisibilitySyncBurst();
      }
    }
  };

  return {
    getModalInputExclusive: (): boolean => modalInputExclusive,
    handleModalInputStateChange,
  };
}
