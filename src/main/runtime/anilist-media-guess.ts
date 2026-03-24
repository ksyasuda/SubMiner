import type { AnilistMediaGuess } from '../../core/services/anilist/anilist-updater';
import { isYoutubeMediaPath } from './youtube-playback';

export type AnilistMediaGuessRuntimeState = {
  mediaKey: string | null;
  mediaDurationSec: number | null;
  mediaGuess: AnilistMediaGuess | null;
  mediaGuessPromise: Promise<AnilistMediaGuess | null> | null;
  lastDurationProbeAtMs: number;
};

type GuessAnilistMediaInfo = (
  mediaPath: string | null,
  mediaTitle: string | null,
) => Promise<AnilistMediaGuess | null>;

export function createMaybeProbeAnilistDurationHandler(deps: {
  getState: () => AnilistMediaGuessRuntimeState;
  setState: (state: AnilistMediaGuessRuntimeState) => void;
  durationRetryIntervalMs: number;
  now: () => number;
  requestMpvDuration: () => Promise<unknown>;
  logWarn: (message: string, error: unknown) => void;
}) {
  return async (mediaKey: string): Promise<number | null> => {
    const state = deps.getState();
    if (state.mediaKey !== mediaKey) {
      return null;
    }
    if (isYoutubeMediaPath(mediaKey)) {
      return null;
    }
    if (typeof state.mediaDurationSec === 'number' && state.mediaDurationSec > 0) {
      return state.mediaDurationSec;
    }
    const now = deps.now();
    if (now - state.lastDurationProbeAtMs < deps.durationRetryIntervalMs) {
      return null;
    }

    deps.setState({
      ...state,
      lastDurationProbeAtMs: now,
    });

    try {
      const durationCandidate = await deps.requestMpvDuration();
      const duration =
        typeof durationCandidate === 'number' && Number.isFinite(durationCandidate)
          ? durationCandidate
          : null;
      const latestState = deps.getState();
      if (duration && duration > 0 && latestState.mediaKey === mediaKey) {
        deps.setState({
          ...latestState,
          mediaDurationSec: duration,
        });
        return duration;
      }
    } catch (error) {
      deps.logWarn('AniList duration probe failed:', error);
    }
    return null;
  };
}

export function createEnsureAnilistMediaGuessHandler(deps: {
  getState: () => AnilistMediaGuessRuntimeState;
  setState: (state: AnilistMediaGuessRuntimeState) => void;
  resolveMediaPathForJimaku: (currentMediaPath: string | null) => string | null;
  getCurrentMediaPath: () => string | null;
  getCurrentMediaTitle: () => string | null;
  guessAnilistMediaInfo: GuessAnilistMediaInfo;
}) {
  return async (mediaKey: string): Promise<AnilistMediaGuess | null> => {
    const state = deps.getState();
    if (state.mediaKey !== mediaKey) {
      return null;
    }
    if (isYoutubeMediaPath(mediaKey)) {
      return null;
    }
    if (state.mediaGuess) {
      return state.mediaGuess;
    }
    if (state.mediaGuessPromise) {
      return state.mediaGuessPromise;
    }

    const mediaPathForGuess = deps.resolveMediaPathForJimaku(deps.getCurrentMediaPath());
    const promise = deps
      .guessAnilistMediaInfo(mediaPathForGuess, deps.getCurrentMediaTitle())
      .then((guess) => {
        const latestState = deps.getState();
        if (latestState.mediaKey === mediaKey) {
          deps.setState({
            ...latestState,
            mediaGuess: guess,
          });
        }
        return guess;
      })
      .finally(() => {
        const latestState = deps.getState();
        if (latestState.mediaKey === mediaKey) {
          deps.setState({
            ...latestState,
            mediaGuessPromise: null,
          });
        }
      });

    deps.setState({
      ...state,
      mediaGuessPromise: promise,
    });
    return promise;
  };
}
