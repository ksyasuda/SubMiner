import type { CliArgs, CliCommandSource } from '../../../cli/args';
import { createCliCommandContextFactory } from '../cli-command-context-factory';
import { createCliCommandRuntimeHandler } from '../cli-command-runtime-handler';
import { createInitialArgsRuntimeHandler } from '../initial-args-runtime-handler';
import type { ComposerInputs, ComposerOutputs } from './contracts';

type CliCommandContextMainDeps = Parameters<typeof createCliCommandContextFactory>[0];
type CliCommandContext = ReturnType<ReturnType<typeof createCliCommandContextFactory>>;
type CliCommandRuntimeHandlerMainDeps = Omit<
  Parameters<typeof createCliCommandRuntimeHandler<CliCommandContext>>[0],
  'createCliCommandContext'
>;
type InitialArgsRuntimeHandlerMainDeps = Omit<
  Parameters<typeof createInitialArgsRuntimeHandler>[0],
  'handleCliCommand'
>;

export type CliStartupComposerOptions = ComposerInputs<{
  cliCommandContextMainDeps: CliCommandContextMainDeps;
  cliCommandRuntimeHandlerMainDeps: CliCommandRuntimeHandlerMainDeps;
  initialArgsRuntimeHandlerMainDeps: InitialArgsRuntimeHandlerMainDeps;
}>;

export type CliStartupComposerResult = ComposerOutputs<{
  createCliCommandContext: () => CliCommandContext;
  handleCliCommand: (args: CliArgs, source?: CliCommandSource) => void;
  handleInitialArgs: () => void;
}>;

export function composeCliStartupHandlers(
  options: CliStartupComposerOptions,
): CliStartupComposerResult {
  const createCliCommandContext = createCliCommandContextFactory(
    options.cliCommandContextMainDeps,
  );
  const handleCliCommand = createCliCommandRuntimeHandler({
    ...options.cliCommandRuntimeHandlerMainDeps,
    createCliCommandContext: () => createCliCommandContext(),
  });
  const handleInitialArgs = createInitialArgsRuntimeHandler({
    ...options.initialArgsRuntimeHandlerMainDeps,
    handleCliCommand: (args, source) => handleCliCommand(args, source),
  });

  return {
    createCliCommandContext,
    handleCliCommand,
    handleInitialArgs,
  };
}
