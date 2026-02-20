type ResolvedConfigLike = {
  immersionTracking?: {
    dbPath?: string | null;
  };
};

type ImmersionTrackerLike = {
  handleMediaChange: (path: string, title: string | null) => void;
};

type MpvClientLike = {
  currentVideoPath?: string | null;
  connected?: boolean;
  requestProperty?: (name: string) => Promise<unknown>;
};

type ImmersionMediaState = {
  path: string | null;
  title: string | null;
};

export type ImmersionMediaRuntimeDeps = {
  getResolvedConfig: () => ResolvedConfigLike;
  defaultImmersionDbPath: string;
  getTracker: () => ImmersionTrackerLike | null;
  getMpvClient: () => MpvClientLike | null;
  getCurrentMediaPath: () => string | null | undefined;
  getCurrentMediaTitle: () => string | null | undefined;
  sleep?: (ms: number) => Promise<void>;
  seedWaitMs?: number;
  seedAttempts?: number;
  logDebug: (message: string) => void;
  logInfo: (message: string) => void;
};

function trimToNull(value: string | null | undefined): string | null {
  if (typeof value !== 'string') return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

async function readMpvPropertyAsString(
  mpvClient: MpvClientLike | null | undefined,
  propertyName: string,
): Promise<string | null> {
  const requestProperty = mpvClient?.requestProperty;
  if (!requestProperty) {
    return null;
  }
  try {
    const value = await requestProperty(propertyName);
    return typeof value === 'string' ? trimToNull(value) : null;
  } catch {
    return null;
  }
}

export function createImmersionMediaRuntime(deps: ImmersionMediaRuntimeDeps): {
  getConfiguredDbPath: () => string;
  seedFromCurrentMedia: () => Promise<void>;
  syncFromCurrentMediaState: () => void;
} {
  const sleep = deps.sleep ?? ((ms: number) => new Promise((resolve) => setTimeout(resolve, ms)));
  const waitMs = deps.seedWaitMs ?? 250;
  const attempts = deps.seedAttempts ?? 120;
  let isSeedInProgress = false;

  const getConfiguredDbPath = (): string => {
    const configuredDbPath = trimToNull(deps.getResolvedConfig().immersionTracking?.dbPath);
    return configuredDbPath ?? deps.defaultImmersionDbPath;
  };

  const getCurrentMpvMediaStateForTracker = async (): Promise<ImmersionMediaState> => {
    const statePath = trimToNull(deps.getCurrentMediaPath());
    const stateTitle = trimToNull(deps.getCurrentMediaTitle());
    if (statePath) {
      return {
        path: statePath,
        title: stateTitle,
      };
    }

    const mpvClient = deps.getMpvClient();
    const trackedPath = trimToNull(mpvClient?.currentVideoPath);
    if (trackedPath) {
      return {
        path: trackedPath,
        title: stateTitle,
      };
    }

    const [pathFromProperty, filenameFromProperty, titleFromProperty] = await Promise.all([
      readMpvPropertyAsString(mpvClient, 'path'),
      readMpvPropertyAsString(mpvClient, 'filename'),
      readMpvPropertyAsString(mpvClient, 'media-title'),
    ]);

    return {
      path: pathFromProperty || filenameFromProperty || null,
      title: stateTitle || titleFromProperty || null,
    };
  };

  const seedFromCurrentMedia = async (): Promise<void> => {
    const tracker = deps.getTracker();
    if (!tracker) {
      deps.logDebug('Immersion tracker seeding skipped: tracker not initialized.');
      return;
    }
    if (isSeedInProgress) {
      deps.logDebug('Immersion tracker seeding already in progress; skipping duplicate call.');
      return;
    }
    deps.logDebug('Starting immersion tracker media-state seed loop.');
    isSeedInProgress = true;

    try {
      for (let attempt = 0; attempt < attempts; attempt += 1) {
        const mediaState = await getCurrentMpvMediaStateForTracker();
        if (mediaState.path) {
          deps.logInfo(
            `Seeded immersion tracker media state at attempt ${attempt + 1}/${attempts}: ${mediaState.path}`,
          );
          tracker.handleMediaChange(mediaState.path, mediaState.title);
          return;
        }

        const mpvClient = deps.getMpvClient();
        if (!mpvClient || !mpvClient.connected) {
          await sleep(waitMs);
          continue;
        }
        if (attempt < attempts - 1) {
          await sleep(waitMs);
        }
      }

      deps.logInfo(
        'Immersion tracker seed failed: media path still unavailable after startup warmup',
      );
    } finally {
      isSeedInProgress = false;
    }
  };

  const syncFromCurrentMediaState = (): void => {
    const tracker = deps.getTracker();
    if (!tracker) {
      deps.logDebug('Immersion tracker sync skipped: tracker not initialized yet.');
      return;
    }

    const pathFromState =
      trimToNull(deps.getCurrentMediaPath()) || trimToNull(deps.getMpvClient()?.currentVideoPath);
    if (pathFromState) {
      deps.logDebug('Immersion tracker sync using path from current media state.');
      tracker.handleMediaChange(pathFromState, trimToNull(deps.getCurrentMediaTitle()));
      return;
    }

    if (!isSeedInProgress) {
      deps.logDebug('Immersion tracker sync did not find media path; starting seed loop.');
      void seedFromCurrentMedia();
    } else {
      deps.logDebug('Immersion tracker sync found seed loop already running.');
    }
  };

  return {
    getConfiguredDbPath,
    seedFromCurrentMedia,
    syncFromCurrentMediaState,
  };
}
