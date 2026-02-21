import {
  createBuildEnsureAnilistMediaGuessMainDepsHandler,
  createBuildGetAnilistMediaGuessRuntimeStateMainDepsHandler,
  createBuildGetCurrentAnilistMediaKeyMainDepsHandler,
  createBuildMaybeProbeAnilistDurationMainDepsHandler,
  createBuildMaybeRunAnilistPostWatchUpdateMainDepsHandler,
  createBuildProcessNextAnilistRetryUpdateMainDepsHandler,
  createBuildRefreshAnilistClientSecretStateMainDepsHandler,
  createBuildResetAnilistMediaGuessStateMainDepsHandler,
  createBuildResetAnilistMediaTrackingMainDepsHandler,
  createBuildSetAnilistMediaGuessRuntimeStateMainDepsHandler,
  createEnsureAnilistMediaGuessHandler,
  createGetAnilistMediaGuessRuntimeStateHandler,
  createGetCurrentAnilistMediaKeyHandler,
  createMaybeProbeAnilistDurationHandler,
  createMaybeRunAnilistPostWatchUpdateHandler,
  createProcessNextAnilistRetryUpdateHandler,
  createRefreshAnilistClientSecretStateHandler,
  createResetAnilistMediaGuessStateHandler,
  createResetAnilistMediaTrackingHandler,
  createSetAnilistMediaGuessRuntimeStateHandler,
} from '../domains/anilist';

export type AnilistTrackingComposerOptions = {
  refreshClientSecretMainDeps: Parameters<
    typeof createBuildRefreshAnilistClientSecretStateMainDepsHandler
  >[0];
  getCurrentMediaKeyMainDeps: Parameters<
    typeof createBuildGetCurrentAnilistMediaKeyMainDepsHandler
  >[0];
  resetMediaTrackingMainDeps: Parameters<
    typeof createBuildResetAnilistMediaTrackingMainDepsHandler
  >[0];
  getMediaGuessRuntimeStateMainDeps: Parameters<
    typeof createBuildGetAnilistMediaGuessRuntimeStateMainDepsHandler
  >[0];
  setMediaGuessRuntimeStateMainDeps: Parameters<
    typeof createBuildSetAnilistMediaGuessRuntimeStateMainDepsHandler
  >[0];
  resetMediaGuessStateMainDeps: Parameters<
    typeof createBuildResetAnilistMediaGuessStateMainDepsHandler
  >[0];
  maybeProbeDurationMainDeps: Parameters<
    typeof createBuildMaybeProbeAnilistDurationMainDepsHandler
  >[0];
  ensureMediaGuessMainDeps: Parameters<typeof createBuildEnsureAnilistMediaGuessMainDepsHandler>[0];
  processNextRetryUpdateMainDeps: Parameters<
    typeof createBuildProcessNextAnilistRetryUpdateMainDepsHandler
  >[0];
  maybeRunPostWatchUpdateMainDeps: Parameters<
    typeof createBuildMaybeRunAnilistPostWatchUpdateMainDepsHandler
  >[0];
};

export type AnilistTrackingComposerResult = {
  refreshAnilistClientSecretState: ReturnType<typeof createRefreshAnilistClientSecretStateHandler>;
  getCurrentAnilistMediaKey: ReturnType<typeof createGetCurrentAnilistMediaKeyHandler>;
  resetAnilistMediaTracking: ReturnType<typeof createResetAnilistMediaTrackingHandler>;
  getAnilistMediaGuessRuntimeState: ReturnType<
    typeof createGetAnilistMediaGuessRuntimeStateHandler
  >;
  setAnilistMediaGuessRuntimeState: ReturnType<
    typeof createSetAnilistMediaGuessRuntimeStateHandler
  >;
  resetAnilistMediaGuessState: ReturnType<typeof createResetAnilistMediaGuessStateHandler>;
  maybeProbeAnilistDuration: ReturnType<typeof createMaybeProbeAnilistDurationHandler>;
  ensureAnilistMediaGuess: ReturnType<typeof createEnsureAnilistMediaGuessHandler>;
  processNextAnilistRetryUpdate: ReturnType<typeof createProcessNextAnilistRetryUpdateHandler>;
  maybeRunAnilistPostWatchUpdate: ReturnType<typeof createMaybeRunAnilistPostWatchUpdateHandler>;
};

export function composeAnilistTrackingHandlers(
  options: AnilistTrackingComposerOptions,
): AnilistTrackingComposerResult {
  const refreshAnilistClientSecretState = createRefreshAnilistClientSecretStateHandler(
    createBuildRefreshAnilistClientSecretStateMainDepsHandler(
      options.refreshClientSecretMainDeps,
    )(),
  );
  const getCurrentAnilistMediaKey = createGetCurrentAnilistMediaKeyHandler(
    createBuildGetCurrentAnilistMediaKeyMainDepsHandler(options.getCurrentMediaKeyMainDeps)(),
  );
  const resetAnilistMediaTracking = createResetAnilistMediaTrackingHandler(
    createBuildResetAnilistMediaTrackingMainDepsHandler(options.resetMediaTrackingMainDeps)(),
  );
  const getAnilistMediaGuessRuntimeState = createGetAnilistMediaGuessRuntimeStateHandler(
    createBuildGetAnilistMediaGuessRuntimeStateMainDepsHandler(
      options.getMediaGuessRuntimeStateMainDeps,
    )(),
  );
  const setAnilistMediaGuessRuntimeState = createSetAnilistMediaGuessRuntimeStateHandler(
    createBuildSetAnilistMediaGuessRuntimeStateMainDepsHandler(
      options.setMediaGuessRuntimeStateMainDeps,
    )(),
  );
  const resetAnilistMediaGuessState = createResetAnilistMediaGuessStateHandler(
    createBuildResetAnilistMediaGuessStateMainDepsHandler(options.resetMediaGuessStateMainDeps)(),
  );
  const maybeProbeAnilistDuration = createMaybeProbeAnilistDurationHandler(
    createBuildMaybeProbeAnilistDurationMainDepsHandler(options.maybeProbeDurationMainDeps)(),
  );
  const ensureAnilistMediaGuess = createEnsureAnilistMediaGuessHandler(
    createBuildEnsureAnilistMediaGuessMainDepsHandler(options.ensureMediaGuessMainDeps)(),
  );
  const processNextAnilistRetryUpdate = createProcessNextAnilistRetryUpdateHandler(
    createBuildProcessNextAnilistRetryUpdateMainDepsHandler(
      options.processNextRetryUpdateMainDeps,
    )(),
  );
  const maybeRunAnilistPostWatchUpdate = createMaybeRunAnilistPostWatchUpdateHandler(
    createBuildMaybeRunAnilistPostWatchUpdateMainDepsHandler(
      options.maybeRunPostWatchUpdateMainDeps,
    )(),
  );

  return {
    refreshAnilistClientSecretState,
    getCurrentAnilistMediaKey,
    resetAnilistMediaTracking,
    getAnilistMediaGuessRuntimeState,
    setAnilistMediaGuessRuntimeState,
    resetAnilistMediaGuessState,
    maybeProbeAnilistDuration,
    ensureAnilistMediaGuess,
    processNextAnilistRetryUpdate,
    maybeRunAnilistPostWatchUpdate,
  };
}
