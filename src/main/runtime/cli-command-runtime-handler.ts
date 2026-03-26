import type { CliArgs, CliCommandSource } from '../../cli/args';
import { createHandleTexthookerOnlyModeTransitionHandler } from './cli-command-prechecks';
import { createBuildHandleTexthookerOnlyModeTransitionMainDepsHandler } from './cli-command-prechecks-main-deps';

type HandleTexthookerOnlyModeTransitionMainDeps = Parameters<
  typeof createBuildHandleTexthookerOnlyModeTransitionMainDepsHandler
>[0];

export function createCliCommandRuntimeHandler<TCliContext>(deps: {
  handleTexthookerOnlyModeTransitionMainDeps: HandleTexthookerOnlyModeTransitionMainDeps;
  createCliCommandContext: () => TCliContext;
  handleCliCommandRuntimeServiceWithContext: (
    args: CliArgs,
    source: CliCommandSource,
    cliContext: TCliContext,
  ) => void;
}) {
  const handleTexthookerOnlyModeTransitionHandler = createHandleTexthookerOnlyModeTransitionHandler(
    createBuildHandleTexthookerOnlyModeTransitionMainDepsHandler(
      deps.handleTexthookerOnlyModeTransitionMainDeps,
    )(),
  );

  return (args: CliArgs, source: CliCommandSource = 'initial'): void => {
    handleTexthookerOnlyModeTransitionHandler(args);
    if (
      !deps.handleTexthookerOnlyModeTransitionMainDeps.isTexthookerOnlyMode() &&
      deps.handleTexthookerOnlyModeTransitionMainDeps.commandNeedsOverlayStartupPrereqs(args)
    ) {
      deps.handleTexthookerOnlyModeTransitionMainDeps.ensureOverlayStartupPrereqs();
    }
    const cliContext = deps.createCliCommandContext();
    deps.handleCliCommandRuntimeServiceWithContext(args, source, cliContext);
  };
}
