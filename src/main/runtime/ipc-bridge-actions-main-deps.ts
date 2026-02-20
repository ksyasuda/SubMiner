import type { createHandleMpvCommandFromIpcHandler } from './ipc-bridge-actions';

type HandleMpvCommandFromIpcMainDeps = Parameters<typeof createHandleMpvCommandFromIpcHandler>[0];

export function createBuildHandleMpvCommandFromIpcMainDepsHandler(
  deps: HandleMpvCommandFromIpcMainDeps,
) {
  return (): HandleMpvCommandFromIpcMainDeps => ({
    handleMpvCommandFromIpcRuntime: (command, options) =>
      deps.handleMpvCommandFromIpcRuntime(command, options),
    buildMpvCommandDeps: () => deps.buildMpvCommandDeps(),
  });
}

export function createBuildRunSubsyncManualFromIpcMainDepsHandler<TRequest, TResult>(deps: {
  runManualFromIpc: (request: TRequest) => Promise<TResult>;
}) {
  return () => ({
    runManualFromIpc: (request: TRequest) => deps.runManualFromIpc(request),
  });
}
