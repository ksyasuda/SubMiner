import type {
  ImmersionTrackerStartupDeps,
  createImmersionTrackerStartupHandler,
} from './immersion-startup';

type ImmersionTrackerStartupMainDeps = Parameters<typeof createImmersionTrackerStartupHandler>[0];

export function createBuildImmersionTrackerStartupMainDepsHandler(
  deps: ImmersionTrackerStartupMainDeps,
) {
  return (): ImmersionTrackerStartupDeps => ({
    getResolvedConfig: () => deps.getResolvedConfig(),
    getConfiguredDbPath: () => deps.getConfiguredDbPath(),
    createTrackerService: (params) => deps.createTrackerService(params),
    setTracker: (tracker) => deps.setTracker(tracker),
    getMpvClient: () => deps.getMpvClient(),
    seedTrackerFromCurrentMedia: () => deps.seedTrackerFromCurrentMedia(),
    logInfo: (message: string) => deps.logInfo(message),
    logDebug: (message: string) => deps.logDebug(message),
    logWarn: (message: string, details: unknown) => deps.logWarn(message, details),
  });
}
