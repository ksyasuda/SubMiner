type ReloadableWebContents = {
  isDestroyed?: () => boolean;
  reload: () => void;
};

type ReloadableOverlayWindow = {
  isDestroyed: () => boolean;
  webContents?: ReloadableWebContents;
};

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

    try {
      webContents.reload();
      reloadCount += 1;
    } catch (error) {
      logWarn?.('Failed to reload overlay window after Yomitan extension load.', error);
    }
  }

  return reloadCount;
}
