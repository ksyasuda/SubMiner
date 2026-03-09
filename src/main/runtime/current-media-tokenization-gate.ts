function normalizeMediaPath(mediaPath: string | null | undefined): string | null {
  if (typeof mediaPath !== 'string') {
    return null;
  }
  const trimmed = mediaPath.trim();
  return trimmed.length > 0 ? trimmed : null;
}

export function createCurrentMediaTokenizationGate(): {
  updateCurrentMediaPath: (mediaPath: string | null | undefined) => void;
  markReady: (mediaPath: string | null | undefined) => void;
  waitUntilReady: (mediaPath: string | null | undefined) => Promise<void>;
} {
  let currentMediaPath: string | null = null;
  let readyMediaPath: string | null = null;
  let pendingMediaPath: string | null = null;
  let pendingPromise: Promise<void> | null = null;
  let resolvePending: (() => void) | null = null;

  const resolvePendingWaiter = (): void => {
    resolvePending?.();
    resolvePending = null;
    pendingPromise = null;
    pendingMediaPath = null;
  };

  const ensurePendingPromise = (mediaPath: string): Promise<void> => {
    if (pendingMediaPath === mediaPath && pendingPromise) {
      return pendingPromise;
    }
    resolvePendingWaiter();
    pendingMediaPath = mediaPath;
    pendingPromise = new Promise<void>((resolve) => {
      resolvePending = resolve;
    });
    return pendingPromise;
  };

  return {
    updateCurrentMediaPath: (mediaPath) => {
      const normalizedPath = normalizeMediaPath(mediaPath);
      if (normalizedPath === currentMediaPath) {
        return;
      }
      currentMediaPath = normalizedPath;
      readyMediaPath = null;
      resolvePendingWaiter();
      if (normalizedPath) {
        ensurePendingPromise(normalizedPath);
      }
    },
    markReady: (mediaPath) => {
      const normalizedPath = normalizeMediaPath(mediaPath);
      if (!normalizedPath) {
        return;
      }
      readyMediaPath = normalizedPath;
      if (pendingMediaPath === normalizedPath) {
        resolvePendingWaiter();
      }
    },
    waitUntilReady: async (mediaPath) => {
      const normalizedPath = normalizeMediaPath(mediaPath) ?? currentMediaPath;
      if (!normalizedPath || readyMediaPath === normalizedPath) {
        return;
      }
      await ensurePendingPromise(normalizedPath);
    },
  };
}
