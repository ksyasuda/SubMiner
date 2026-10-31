import {
  createHandleMpvCommandFromIpcHandler,
  createRunSubsyncManualFromIpcHandler,
} from './ipc-bridge-actions';

export function createIpcRuntimeHandlers<TRequest, TResult>(deps: {
  handleMpvCommandFromIpcDeps: Parameters<typeof createHandleMpvCommandFromIpcHandler>[0];
  runSubsyncManualFromIpcDeps: Parameters<
    typeof createRunSubsyncManualFromIpcHandler<TRequest, TResult>
  >[0];
}) {
  const handleMpvCommandFromIpc = createHandleMpvCommandFromIpcHandler(
    deps.handleMpvCommandFromIpcDeps,
  );
  const runSubsyncManualFromIpc = createRunSubsyncManualFromIpcHandler<TRequest, TResult>(
    deps.runSubsyncManualFromIpcDeps,
  );

  return {
    handleMpvCommandFromIpc,
    runSubsyncManualFromIpc,
  };
}
