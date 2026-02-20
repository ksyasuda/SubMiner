import type { MpvCommandFromIpcRuntimeDeps } from '../ipc-mpv-command';

export function createHandleMpvCommandFromIpcHandler(deps: {
  handleMpvCommandFromIpcRuntime: (
    command: (string | number)[],
    options: MpvCommandFromIpcRuntimeDeps,
  ) => void;
  buildMpvCommandDeps: () => MpvCommandFromIpcRuntimeDeps;
}) {
  return (command: (string | number)[]): void => {
    deps.handleMpvCommandFromIpcRuntime(command, deps.buildMpvCommandDeps());
  };
}

export function createRunSubsyncManualFromIpcHandler<TRequest, TResult>(deps: {
  runManualFromIpc: (request: TRequest) => Promise<TResult>;
}) {
  return async (request: TRequest): Promise<TResult> => {
    return deps.runManualFromIpc(request);
  };
}
