import type { CliArgs } from '../../cli/args';
import {
  isHeadlessInitialCommand,
  isStandaloneTexthookerCommand,
  shouldRunSettingsOnlyStartup,
} from '../../cli/args';

export function getStartupModeFlags(initialArgs: CliArgs | null | undefined): {
  shouldUseMinimalStartup: boolean;
  shouldSkipHeavyStartup: boolean;
} {
  return {
    shouldUseMinimalStartup: Boolean(
      (initialArgs && isStandaloneTexthookerCommand(initialArgs)) ||
      initialArgs?.configSettings ||
      initialArgs?.update ||
      (initialArgs?.stats &&
        (initialArgs.statsCleanup || initialArgs.statsBackground || initialArgs.statsStop)),
    ),
    shouldSkipHeavyStartup: Boolean(
      initialArgs &&
      (shouldRunSettingsOnlyStartup(initialArgs) ||
        initialArgs.configSettings ||
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
  return !(initialArgs && (isHeadlessInitialCommand(initialArgs) || initialArgs.configSettings));
}

export function shouldStartAutomaticUpdateChecks(initialArgs: CliArgs | null | undefined): boolean {
  return !(initialArgs && (isHeadlessInitialCommand(initialArgs) || initialArgs.configSettings));
}
