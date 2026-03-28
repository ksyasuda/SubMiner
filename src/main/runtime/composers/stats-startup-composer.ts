import type { ComposerInputs, ComposerOutputs } from './contracts';

type BackgroundStatsStartResult = {
  url: string;
  runningInCurrentProcess: boolean;
};

type BackgroundStatsStopResult = {
  ok: boolean;
  stale: boolean;
};

export type StatsStartupComposerOptions = ComposerInputs<{
  ensureStatsServerStarted: () => string;
  ensureBackgroundStatsServerStarted: () => BackgroundStatsStartResult;
  stopBackgroundStatsServer: () => Promise<BackgroundStatsStopResult> | BackgroundStatsStopResult;
  ensureImmersionTrackerStarted: () => void;
}>;

export type StatsStartupComposerResult = ComposerOutputs<StatsStartupComposerOptions>;

export function composeStatsStartupRuntime(
  options: StatsStartupComposerOptions,
): StatsStartupComposerResult {
  return options;
}
