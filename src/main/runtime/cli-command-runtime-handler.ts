import type { CliArgs, CliCommandSource } from '../../cli/args';
import { createHandleTexthookerOnlyModeTransitionHandler } from './cli-command-prechecks';

type HandleTexthookerOnlyModeTransitionMainDeps = Parameters<
  typeof createHandleTexthookerOnlyModeTransitionHandler
>[0];

export function createCliCommandRuntimeHandler<TCliContext>(deps: {
  handleTexthookerOnlyModeTransitionMainDeps: HandleTexthookerOnlyModeTransitionMainDeps;
  ensureTrayForCommand?: (args: CliArgs, source: CliCommandSource) => void;
  createCliCommandContext: () => TCliContext;
  handleCliCommandRuntimeServiceWithContext: (
    args: CliArgs,
    source: CliCommandSource,
    cliContext: TCliContext,
  ) => void;
}) {
  const handleTexthookerOnlyModeTransitionHandler = createHandleTexthookerOnlyModeTransitionHandler(
    deps.handleTexthookerOnlyModeTransitionMainDeps,
  );

  return (args: CliArgs, source: CliCommandSource = 'initial'): void => {
    handleTexthookerOnlyModeTransitionHandler(args);
    if (
      !deps.handleTexthookerOnlyModeTransitionMainDeps.isTexthookerOnlyMode() &&
      deps.handleTexthookerOnlyModeTransitionMainDeps.commandNeedsOverlayStartupPrereqs(args)
    ) {
      deps.handleTexthookerOnlyModeTransitionMainDeps.ensureOverlayStartupPrereqs();
    }
    deps.ensureTrayForCommand?.(args, source);
    const cliContext = deps.createCliCommandContext();
    deps.handleCliCommandRuntimeServiceWithContext(args, source, cliContext);
  };
}
