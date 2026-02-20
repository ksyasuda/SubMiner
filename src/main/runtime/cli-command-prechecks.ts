import type { CliArgs } from '../../cli/args';

export function createHandleTexthookerOnlyModeTransitionHandler(deps: {
  isTexthookerOnlyMode: () => boolean;
  setTexthookerOnlyMode: (enabled: boolean) => void;
  commandNeedsOverlayRuntime: (args: CliArgs) => boolean;
  startBackgroundWarmups: () => void;
  logInfo: (message: string) => void;
}) {
  return (args: CliArgs): void => {
    if (
      deps.isTexthookerOnlyMode() &&
      !args.texthooker &&
      (args.start || deps.commandNeedsOverlayRuntime(args))
    ) {
      deps.setTexthookerOnlyMode(false);
      deps.logInfo('Disabling texthooker-only mode after overlay/start command.');
      deps.startBackgroundWarmups();
    }
  };
}
