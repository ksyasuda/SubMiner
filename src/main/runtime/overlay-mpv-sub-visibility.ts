type MpvVisibilityClient = {
  connected: boolean;
  requestProperty: (name: string) => Promise<unknown>;
};

function parseSubVisibility(value: unknown): boolean {
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (normalized === 'no' || normalized === 'false' || normalized === '0') {
      return false;
    }
    if (normalized === 'yes' || normalized === 'true' || normalized === '1') {
      return true;
    }
  }

  if (typeof value === 'boolean') {
    return value;
  }

  if (typeof value === 'number') {
    return value !== 0;
  }

  return true;
}

export function createEnsureOverlayMpvSubtitlesHiddenHandler(deps: {
  getMpvClient: () => MpvVisibilityClient | null;
  getSavedSubVisibility: () => boolean | null;
  setSavedSubVisibility: (visible: boolean | null) => void;
  getRevision: () => number;
  setRevision: (revision: number) => void;
  setMpvSubVisibility: (visible: boolean) => void;
  logWarn: (message: string, error: unknown) => void;
}) {
  return async (): Promise<void> => {
    const revision = deps.getRevision() + 1;
    deps.setRevision(revision);

    const mpvClient = deps.getMpvClient();
    if (!mpvClient || !mpvClient.connected) {
      return;
    }

    const shouldCaptureSavedVisibility = deps.getSavedSubVisibility() === null;
    const savedVisibilityPromise = shouldCaptureSavedVisibility
      ? mpvClient.requestProperty('sub-visibility')
      : null;

    // Hide immediately on overlay toggle; capture/restore logic is handled separately.
    deps.setMpvSubVisibility(false);

    if (shouldCaptureSavedVisibility && savedVisibilityPromise) {
      try {
        const currentSubVisibility = await savedVisibilityPromise;
        if (revision !== deps.getRevision()) {
          return;
        }
        deps.setSavedSubVisibility(parseSubVisibility(currentSubVisibility));
      } catch (error) {
        if (revision !== deps.getRevision()) {
          return;
        }
        deps.logWarn(
          '[overlay] Failed to capture mpv sub-visibility; falling back to visible restore',
          error,
        );
        deps.setSavedSubVisibility(true);
      }
    }
  };
}

export function createRestoreOverlayMpvSubtitlesHandler(deps: {
  getSavedSubVisibility: () => boolean | null;
  setSavedSubVisibility: (visible: boolean | null) => void;
  getRevision: () => number;
  setRevision: (revision: number) => void;
  isMpvConnected: () => boolean;
  shouldKeepSuppressedFromVisibleOverlayBinding: () => boolean;
  setMpvSubVisibility: (visible: boolean) => void;
}) {
  return (): void => {
    deps.setRevision(deps.getRevision() + 1);

    const savedVisibility = deps.getSavedSubVisibility();
    if (deps.shouldKeepSuppressedFromVisibleOverlayBinding()) {
      deps.setMpvSubVisibility(false);
      return;
    }

    if (savedVisibility === null) {
      return;
    }

    if (!deps.isMpvConnected()) {
      return;
    }

    if (savedVisibility !== null) {
      deps.setMpvSubVisibility(savedVisibility);
    }

    deps.setSavedSubVisibility(null);
  };
}
