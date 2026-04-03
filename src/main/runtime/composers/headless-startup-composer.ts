import { createStartupRuntimeHandlers } from '../startup-runtime-handlers';
import type { ComposerInputs, ComposerOutputs } from './contracts';

type StartupRuntimeHandlersDeps<TCliArgs, TStartupState, TStartupBootstrapRuntimeDeps> = Parameters<
  typeof createStartupRuntimeHandlers<TCliArgs, TStartupState, TStartupBootstrapRuntimeDeps>
>[0];
type StartupRuntimeHandlers<TCliArgs, TStartupState, TStartupBootstrapRuntimeDeps> = ReturnType<
  typeof createStartupRuntimeHandlers<TCliArgs, TStartupState, TStartupBootstrapRuntimeDeps>
>;

export type HeadlessStartupComposerOptions<TCliArgs, TStartupState, TStartupBootstrapRuntimeDeps> =
  ComposerInputs<{
    startupRuntimeHandlersDeps: StartupRuntimeHandlersDeps<
      TCliArgs,
      TStartupState,
      TStartupBootstrapRuntimeDeps
    >;
  }>;

export type HeadlessStartupComposerResult<TCliArgs, TStartupState, TStartupBootstrapRuntimeDeps> =
  ComposerOutputs<
    Pick<
      StartupRuntimeHandlers<TCliArgs, TStartupState, TStartupBootstrapRuntimeDeps>,
      'appLifecycleRuntimeRunner' | 'runAndApplyStartupState'
    >
  >;

export function composeHeadlessStartupHandlers<
  TCliArgs,
  TStartupState,
  TStartupBootstrapRuntimeDeps,
>(
  options: HeadlessStartupComposerOptions<TCliArgs, TStartupState, TStartupBootstrapRuntimeDeps>,
): HeadlessStartupComposerResult<TCliArgs, TStartupState, TStartupBootstrapRuntimeDeps> {
  const { appLifecycleRuntimeRunner, runAndApplyStartupState } = createStartupRuntimeHandlers(
    options.startupRuntimeHandlersDeps,
  );

  return {
    appLifecycleRuntimeRunner,
    runAndApplyStartupState,
  };
}
