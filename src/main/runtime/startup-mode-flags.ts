import type { CliArgs } from '../../cli/args';
import {
  isHeadlessInitialCommand,
  isStandaloneTexthookerCommand,
  shouldRunYomitanOnlyStartup,
} from '../../cli/args';

export function getStartupModeFlags(initialArgs: CliArgs | null | undefined): {
  shouldUseMinimalStartup: boolean;
  shouldSkipHeavyStartup: boolean;
} {
  return {
    shouldUseMinimalStartup: Boolean(
      (initialArgs && isStandaloneTexthookerCommand(initialArgs)) ||
      initialArgs?.settings ||
      initialArgs?.update ||
      (initialArgs?.stats &&
        (initialArgs.statsCleanup || initialArgs.statsBackground || initialArgs.statsStop)),
    ),
    shouldSkipHeavyStartup: Boolean(
      initialArgs &&
      (shouldRunYomitanOnlyStartup(initialArgs) ||
        initialArgs.settings ||
        initialArgs.stats ||
        initialArgs.dictionary ||
        initialArgs.update ||
        initialArgs.setup),
    ),
  };
}

export function shouldRefreshAnilistOnConfigReload(
  initialArgs: CliArgs | null | undefined,
): boolean {
  return !(initialArgs && (isHeadlessInitialCommand(initialArgs) || initialArgs.settings));
}

export function shouldStartAutomaticUpdateChecks(initialArgs: CliArgs | null | undefined): boolean {
  return !(initialArgs && (isHeadlessInitialCommand(initialArgs) || initialArgs.settings));
}
