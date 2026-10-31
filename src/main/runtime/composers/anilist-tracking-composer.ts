import {
  createEnsureAnilistMediaGuessHandler,
  createGetAnilistMediaGuessRuntimeStateHandler,
  createGetCurrentAnilistMediaKeyHandler,
  createMaybeProbeAnilistDurationHandler,
  createMaybeRunAnilistPostWatchUpdateHandler,
  createProcessNextAnilistRetryUpdateHandler,
  createRecordAnilistMediaDurationHandler,
  createRefreshAnilistClientSecretStateHandler,
  createResetAnilistMediaGuessStateHandler,
  createResetAnilistMediaTrackingHandler,
  createSetAnilistMediaGuessRuntimeStateHandler,
} from '../domains/anilist';
import type { ComposerInputs, ComposerOutputs } from './contracts';

export type AnilistTrackingComposerOptions = ComposerInputs<{
  refreshClientSecretMainDeps: Parameters<typeof createRefreshAnilistClientSecretStateHandler>[0];
  getCurrentMediaKeyMainDeps: Parameters<typeof createGetCurrentAnilistMediaKeyHandler>[0];
  resetMediaTrackingMainDeps: Parameters<typeof createResetAnilistMediaTrackingHandler>[0];
  getMediaGuessRuntimeStateMainDeps: Parameters<
    typeof createGetAnilistMediaGuessRuntimeStateHandler
  >[0];
  setMediaGuessRuntimeStateMainDeps: Parameters<
    typeof createSetAnilistMediaGuessRuntimeStateHandler
  >[0];
  recordMediaDurationMainDeps: Parameters<typeof createRecordAnilistMediaDurationHandler>[0];
  resetMediaGuessStateMainDeps: Parameters<typeof createResetAnilistMediaGuessStateHandler>[0];
  maybeProbeDurationMainDeps: Parameters<typeof createMaybeProbeAnilistDurationHandler>[0];
  ensureMediaGuessMainDeps: Parameters<typeof createEnsureAnilistMediaGuessHandler>[0];
  processNextRetryUpdateMainDeps: Parameters<typeof createProcessNextAnilistRetryUpdateHandler>[0];
  maybeRunPostWatchUpdateMainDeps: Parameters<
    typeof createMaybeRunAnilistPostWatchUpdateHandler
  >[0];
}>;

export type AnilistTrackingComposerResult = ComposerOutputs<{
  refreshAnilistClientSecretState: ReturnType<typeof createRefreshAnilistClientSecretStateHandler>;
  getCurrentAnilistMediaKey: ReturnType<typeof createGetCurrentAnilistMediaKeyHandler>;
  resetAnilistMediaTracking: ReturnType<typeof createResetAnilistMediaTrackingHandler>;
  getAnilistMediaGuessRuntimeState: ReturnType<
    typeof createGetAnilistMediaGuessRuntimeStateHandler
  >;
  setAnilistMediaGuessRuntimeState: ReturnType<
    typeof createSetAnilistMediaGuessRuntimeStateHandler
  >;
  recordAnilistMediaDuration: ReturnType<typeof createRecordAnilistMediaDurationHandler>;
  resetAnilistMediaGuessState: ReturnType<typeof createResetAnilistMediaGuessStateHandler>;
  maybeProbeAnilistDuration: ReturnType<typeof createMaybeProbeAnilistDurationHandler>;
  ensureAnilistMediaGuess: ReturnType<typeof createEnsureAnilistMediaGuessHandler>;
  processNextAnilistRetryUpdate: ReturnType<typeof createProcessNextAnilistRetryUpdateHandler>;
  maybeRunAnilistPostWatchUpdate: ReturnType<typeof createMaybeRunAnilistPostWatchUpdateHandler>;
}>;

export function composeAnilistTrackingHandlers(
  options: AnilistTrackingComposerOptions,
): AnilistTrackingComposerResult {
  const refreshAnilistClientSecretState = createRefreshAnilistClientSecretStateHandler(
    options.refreshClientSecretMainDeps,
  );
  const getCurrentAnilistMediaKey = createGetCurrentAnilistMediaKeyHandler(
    options.getCurrentMediaKeyMainDeps,
  );
  const resetAnilistMediaTracking = createResetAnilistMediaTrackingHandler(
    options.resetMediaTrackingMainDeps,
  );
  const getAnilistMediaGuessRuntimeState = createGetAnilistMediaGuessRuntimeStateHandler(
    options.getMediaGuessRuntimeStateMainDeps,
  );
  const setAnilistMediaGuessRuntimeState = createSetAnilistMediaGuessRuntimeStateHandler(
    options.setMediaGuessRuntimeStateMainDeps,
  );
  const recordAnilistMediaDuration = createRecordAnilistMediaDurationHandler(
    options.recordMediaDurationMainDeps,
  );
  const resetAnilistMediaGuessState = createResetAnilistMediaGuessStateHandler(
    options.resetMediaGuessStateMainDeps,
  );
  const maybeProbeAnilistDuration = createMaybeProbeAnilistDurationHandler(
    options.maybeProbeDurationMainDeps,
  );
  const ensureAnilistMediaGuess = createEnsureAnilistMediaGuessHandler(
    options.ensureMediaGuessMainDeps,
  );
  const processNextAnilistRetryUpdate = createProcessNextAnilistRetryUpdateHandler(
    options.processNextRetryUpdateMainDeps,
  );
  const maybeRunAnilistPostWatchUpdate = createMaybeRunAnilistPostWatchUpdateHandler(
    options.maybeRunPostWatchUpdateMainDeps,
  );

  return {
    refreshAnilistClientSecretState,
    getCurrentAnilistMediaKey,
    resetAnilistMediaTracking,
    getAnilistMediaGuessRuntimeState,
    setAnilistMediaGuessRuntimeState,
    recordAnilistMediaDuration,
    resetAnilistMediaGuessState,
    maybeProbeAnilistDuration,
    ensureAnilistMediaGuess,
    processNextAnilistRetryUpdate,
    maybeRunAnilistPostWatchUpdate,
  };
}
