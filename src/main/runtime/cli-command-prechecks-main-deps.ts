import type { CliArgs } from '../../cli/args';

export function createBuildHandleTexthookerOnlyModeTransitionMainDepsHandler(deps: {
  isTexthookerOnlyMode: () => boolean;
  setTexthookerOnlyMode: (enabled: boolean) => void;
  commandNeedsOverlayRuntime: (args: CliArgs) => boolean;
  startBackgroundWarmups: () => void;
  logInfo: (message: string) => void;
}) {
  return () => ({
    isTexthookerOnlyMode: () => deps.isTexthookerOnlyMode(),
    setTexthookerOnlyMode: (enabled: boolean) => deps.setTexthookerOnlyMode(enabled),
    commandNeedsOverlayRuntime: (args: CliArgs) => deps.commandNeedsOverlayRuntime(args),
    startBackgroundWarmups: () => deps.startBackgroundWarmups(),
    logInfo: (message: string) => deps.logInfo(message),
  });
}
