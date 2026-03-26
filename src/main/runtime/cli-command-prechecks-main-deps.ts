import type { CliArgs } from '../../cli/args';

export function createBuildHandleTexthookerOnlyModeTransitionMainDepsHandler(deps: {
  isTexthookerOnlyMode: () => boolean;
  setTexthookerOnlyMode: (enabled: boolean) => void;
  commandNeedsOverlayStartupPrereqs: (args: CliArgs) => boolean;
  ensureOverlayStartupPrereqs: () => void;
  startBackgroundWarmups: () => void;
  logInfo: (message: string) => void;
}) {
  return () => ({
    isTexthookerOnlyMode: () => deps.isTexthookerOnlyMode(),
    setTexthookerOnlyMode: (enabled: boolean) => deps.setTexthookerOnlyMode(enabled),
    commandNeedsOverlayStartupPrereqs: (args: CliArgs) =>
      deps.commandNeedsOverlayStartupPrereqs(args),
    ensureOverlayStartupPrereqs: () => deps.ensureOverlayStartupPrereqs(),
    startBackgroundWarmups: () => deps.startBackgroundWarmups(),
    logInfo: (message: string) => deps.logInfo(message),
  });
}
