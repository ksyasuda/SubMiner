type ReloadableWebContents = {
  isDestroyed?: () => boolean;
  isLoading?: () => boolean;
  once?: (event: 'did-finish-load', listener: () => void) => void;
  reload: () => void;
};

type ReloadableOverlayWindow = {
  isDestroyed: () => boolean;
  webContents?: ReloadableWebContents;
};

function reloadWebContentsForYomitanContentScripts(
  webContents: ReloadableWebContents,
  logWarn?: (message: string, error: unknown) => void,
): boolean {
  try {
    webContents.reload();
    return true;
  } catch (error) {
    logWarn?.('Failed to reload overlay window after Yomitan extension load.', error);
    return false;
  }
}

export function reloadOverlayWindowsForYomitanContentScripts(
  windows: ReloadableOverlayWindow[],
  logWarn?: (message: string, error: unknown) => void,
): number {
  let reloadCount = 0;

  for (const window of windows) {
    if (window.isDestroyed()) {
      continue;
    }

    const webContents = window.webContents;
    if (!webContents || webContents.isDestroyed?.()) {
      continue;
    }
    if (webContents.isLoading?.()) {
      webContents.once?.('did-finish-load', () => {
        if (window.isDestroyed() || webContents.isDestroyed?.()) {
          return;
        }
        if (reloadWebContentsForYomitanContentScripts(webContents, logWarn)) {
          reloadCount += 1;
        }
      });
      continue;
    }

    if (reloadWebContentsForYomitanContentScripts(webContents, logWarn)) {
      reloadCount += 1;
    }
  }

  return reloadCount;
}
