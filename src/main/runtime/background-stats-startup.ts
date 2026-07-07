export interface EnsureBackgroundStatsServerDeps {
  isStatsAutoStartEnabled: () => boolean;
  isImmersionTrackingEnabled: () => boolean;
  ensureBackgroundStatsServerStarted: () => {
    url: string;
    runningInCurrentProcess: boolean;
  };
  logInfo: (message: string) => void;
  logWarn: (message: string, error?: unknown) => void;
}

export function createEnsureBackgroundStatsServerHandler(
  deps: EnsureBackgroundStatsServerDeps,
): () => void {
  return () => {
    if (!deps.isStatsAutoStartEnabled()) {
      deps.logInfo('Background start: stats.autoStartServer is disabled; skipping stats server.');
      return;
    }
    if (!deps.isImmersionTrackingEnabled()) {
      deps.logInfo('Background start: immersion tracking is disabled; skipping stats server.');
      return;
    }
    try {
      const result = deps.ensureBackgroundStatsServerStarted();
      deps.logInfo(
        result.runningInCurrentProcess
          ? `Background start: stats server started at ${result.url}.`
          : `Background start: stats server already running at ${result.url}; skipping.`,
      );
    } catch (error) {
      deps.logWarn('Background start: failed to start stats server.', error);
    }
  };
}
