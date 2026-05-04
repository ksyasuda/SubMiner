import { isYoutubeMediaPath } from './youtube-playback';

type AnilistGuess = {
  title: string;
  episode: number | null;
};

type AnilistUpdateResult = {
  status: 'updated' | 'skipped' | 'error';
  message: string;
};

type RetryQueueItem = {
  key: string;
  title: string;
  episode: number;
};

type AnilistPostWatchRunOptions = {
  force?: boolean;
};

export function buildAnilistAttemptKey(mediaKey: string, episode: number): string {
  return `${mediaKey}::${episode}`;
}

export function rememberAnilistAttemptedUpdateKey(
  attemptedKeys: Set<string>,
  key: string,
  maxSize: number,
): void {
  attemptedKeys.add(key);
  if (attemptedKeys.size <= maxSize) {
    return;
  }
  const oldestKey = attemptedKeys.values().next().value;
  if (typeof oldestKey === 'string') {
    attemptedKeys.delete(oldestKey);
  }
}

export function createProcessNextAnilistRetryUpdateHandler(deps: {
  nextReady: () => RetryQueueItem | null;
  refreshRetryQueueState: () => void;
  setLastAttemptAt: (value: number) => void;
  setLastError: (value: string | null) => void;
  refreshAnilistClientSecretState: () => Promise<string | null>;
  updateAnilistPostWatchProgress: (
    accessToken: string,
    title: string,
    episode: number,
  ) => Promise<AnilistUpdateResult>;
  markSuccess: (key: string) => void;
  rememberAttemptedUpdateKey: (key: string) => void;
  markFailure: (key: string, message: string) => void;
  logInfo: (message: string) => void;
  now: () => number;
}) {
  return async (): Promise<{ ok: boolean; message: string }> => {
    const queued = deps.nextReady();
    deps.refreshRetryQueueState();
    if (!queued) {
      return { ok: true, message: 'AniList queue has no ready items.' };
    }

    deps.setLastAttemptAt(deps.now());
    const accessToken = await deps.refreshAnilistClientSecretState();
    if (!accessToken) {
      deps.setLastError('AniList token unavailable for queued retry.');
      return { ok: false, message: 'AniList token unavailable for queued retry.' };
    }

    const result = await deps.updateAnilistPostWatchProgress(
      accessToken,
      queued.title,
      queued.episode,
    );
    if (result.status === 'updated' || result.status === 'skipped') {
      deps.markSuccess(queued.key);
      deps.rememberAttemptedUpdateKey(queued.key);
      deps.setLastError(null);
      deps.refreshRetryQueueState();
      deps.logInfo(`[AniList queue] ${result.message}`);
      return { ok: true, message: result.message };
    }

    deps.markFailure(queued.key, result.message);
    deps.setLastError(result.message);
    deps.refreshRetryQueueState();
    return { ok: false, message: result.message };
  };
}

export function createMaybeRunAnilistPostWatchUpdateHandler(deps: {
  getInFlight: () => boolean;
  setInFlight: (value: boolean) => void;
  getResolvedConfig: () => unknown;
  isAnilistTrackingEnabled: (config: unknown) => boolean;
  getCurrentMediaKey: () => string | null;
  hasMpvClient: () => boolean;
  getTrackedMediaKey: () => string | null;
  resetTrackedMedia: (mediaKey: string | null) => void;
  getWatchedSeconds: () => number;
  maybeProbeAnilistDuration: (mediaKey: string) => Promise<number | null>;
  ensureAnilistMediaGuess: (mediaKey: string) => Promise<AnilistGuess | null>;
  hasAttemptedUpdateKey: (key: string) => boolean;
  processNextAnilistRetryUpdate: () => Promise<{ ok: boolean; message: string }>;
  refreshAnilistClientSecretState: () => Promise<string | null>;
  enqueueRetry: (key: string, title: string, episode: number) => void;
  markRetryFailure: (key: string, message: string) => void;
  markRetrySuccess: (key: string) => void;
  refreshRetryQueueState: () => void;
  updateAnilistPostWatchProgress: (
    accessToken: string,
    title: string,
    episode: number,
  ) => Promise<AnilistUpdateResult>;
  rememberAttemptedUpdateKey: (key: string) => void;
  showMpvOsd: (message: string) => void;
  logInfo: (message: string) => void;
  logWarn: (message: string) => void;
  minWatchSeconds: number;
  minWatchRatio: number;
}) {
  return async (options: AnilistPostWatchRunOptions = {}): Promise<void> => {
    if (deps.getInFlight()) {
      return;
    }
    const force = options.force === true;

    const resolved = deps.getResolvedConfig();
    if (!deps.isAnilistTrackingEnabled(resolved)) {
      return;
    }

    const mediaKey = deps.getCurrentMediaKey();
    if (!mediaKey || (!force && !deps.hasMpvClient())) {
      return;
    }
    if (isYoutubeMediaPath(mediaKey)) {
      return;
    }
    if (deps.getTrackedMediaKey() !== mediaKey) {
      deps.resetTrackedMedia(mediaKey);
    }

    let watchedSeconds = 0;
    if (!force) {
      watchedSeconds = deps.getWatchedSeconds();
      if (!Number.isFinite(watchedSeconds) || watchedSeconds < deps.minWatchSeconds) {
        return;
      }
    }

    deps.setInFlight(true);
    try {
      if (!force) {
        const duration = await deps.maybeProbeAnilistDuration(mediaKey);
        if (!duration || duration <= 0) {
          return;
        }
        if (watchedSeconds / duration < deps.minWatchRatio) {
          return;
        }
      }

      const guess = await deps.ensureAnilistMediaGuess(mediaKey);
      if (!guess?.title || !guess.episode || guess.episode <= 0) {
        return;
      }

      const attemptKey = buildAnilistAttemptKey(mediaKey, guess.episode);
      if (deps.hasAttemptedUpdateKey(attemptKey)) {
        return;
      }

      await deps.processNextAnilistRetryUpdate();
      if (deps.hasAttemptedUpdateKey(attemptKey)) {
        return;
      }

      const accessToken = await deps.refreshAnilistClientSecretState();
      if (!accessToken) {
        deps.enqueueRetry(attemptKey, guess.title, guess.episode);
        deps.markRetryFailure(attemptKey, 'cannot authenticate without anilist.accessToken');
        deps.refreshRetryQueueState();
        deps.showMpvOsd('AniList: access token not configured');
        return;
      }

      const result = await deps.updateAnilistPostWatchProgress(
        accessToken,
        guess.title,
        guess.episode,
      );
      if (result.status === 'updated') {
        deps.rememberAttemptedUpdateKey(attemptKey);
        deps.markRetrySuccess(attemptKey);
        deps.refreshRetryQueueState();
        deps.showMpvOsd(result.message);
        deps.logInfo(result.message);
        return;
      }
      if (result.status === 'skipped') {
        deps.rememberAttemptedUpdateKey(attemptKey);
        deps.markRetrySuccess(attemptKey);
        deps.refreshRetryQueueState();
        deps.logInfo(result.message);
        return;
      }

      deps.enqueueRetry(attemptKey, guess.title, guess.episode);
      deps.markRetryFailure(attemptKey, result.message);
      deps.refreshRetryQueueState();
      deps.showMpvOsd(`AniList: ${result.message}`);
      deps.logWarn(result.message);
    } finally {
      deps.setInFlight(false);
    }
  };
}
