type MpvVisibilityClient = {
  connected: boolean;
  requestProperty: (name: string) => Promise<unknown>;
};

type RestoreOptions = {
  respectVisibleOverlayBinding?: boolean;
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
  getSavedSecondarySubVisibility: () => boolean | null;
  setSavedSecondarySubVisibility: (visible: boolean | null) => void;
  getRevision: () => number;
  setRevision: (revision: number) => void;
  setMpvSubVisibility: (visible: boolean) => void;
  setMpvSecondarySubVisibility: (visible: boolean) => void;
  logWarn: (message: string, error: unknown) => void;
}) {
  return async (): Promise<void> => {
    const revision = deps.getRevision() + 1;
    deps.setRevision(revision);

    const mpvClient = deps.getMpvClient();
    if (!mpvClient || !mpvClient.connected) {
      return;
    }

    if (deps.getSavedSubVisibility() === null) {
      try {
        const currentSubVisibility = await mpvClient.requestProperty('sub-visibility');
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

    if (deps.getSavedSecondarySubVisibility() === null) {
      try {
        const currentSecondarySubVisibility = await mpvClient.requestProperty('secondary-sub-visibility');
        if (revision !== deps.getRevision()) {
          return;
        }
        deps.setSavedSecondarySubVisibility(parseSubVisibility(currentSecondarySubVisibility));
      } catch (error) {
        if (revision !== deps.getRevision()) {
          return;
        }
        deps.logWarn(
          '[overlay] Failed to capture secondary mpv sub-visibility; falling back to visible restore',
          error,
        );
        deps.setSavedSecondarySubVisibility(true);
      }
    }

    if (revision !== deps.getRevision()) {
      return;
    }

    deps.setMpvSubVisibility(false);
    deps.setMpvSecondarySubVisibility(false);
  };
}

export function createRestoreOverlayMpvSubtitlesHandler(deps: {
  getSavedSubVisibility: () => boolean | null;
  setSavedSubVisibility: (visible: boolean | null) => void;
  getSavedSecondarySubVisibility: () => boolean | null;
  setSavedSecondarySubVisibility: (visible: boolean | null) => void;
  getRevision: () => number;
  setRevision: (revision: number) => void;
  isMpvConnected: () => boolean;
  shouldKeepSuppressedFromVisibleOverlayBinding: () => boolean;
  setMpvSubVisibility: (visible: boolean) => void;
  setMpvSecondarySubVisibility: (visible: boolean) => void;
}) {
  return (options?: RestoreOptions): void => {
    deps.setRevision(deps.getRevision() + 1);

    const savedVisibility = deps.getSavedSubVisibility();
    const respectVisibleOverlayBinding = options?.respectVisibleOverlayBinding ?? true;
    if (
      respectVisibleOverlayBinding &&
      deps.shouldKeepSuppressedFromVisibleOverlayBinding()
    ) {
      deps.setMpvSubVisibility(false);
      deps.setMpvSecondarySubVisibility(false);
      return;
    }

    const hasSecondarySavedVisibility = deps.getSavedSecondarySubVisibility() !== null;

    if (savedVisibility === null && !hasSecondarySavedVisibility) {
      return;
    }

    if (!deps.isMpvConnected()) {
      return;
    }

    if (savedVisibility !== null) {
      deps.setMpvSubVisibility(savedVisibility);
    }
    const savedSecondaryVisibility = deps.getSavedSecondarySubVisibility();
    if (savedSecondaryVisibility !== null) {
      deps.setMpvSecondarySubVisibility(savedSecondaryVisibility);
    }

    deps.setSavedSubVisibility(null);
    deps.setSavedSecondarySubVisibility(null);
  };
}
