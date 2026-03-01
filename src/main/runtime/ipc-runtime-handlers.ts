import {
  createHandleMpvCommandFromIpcHandler,
  createRunSubsyncManualFromIpcHandler,
} from './ipc-bridge-actions';
import {
  createBuildHandleMpvCommandFromIpcMainDepsHandler,
  createBuildRunSubsyncManualFromIpcMainDepsHandler,
} from './ipc-bridge-actions-main-deps';

type HandleMpvCommandFromIpcMainDeps = Parameters<
  typeof createBuildHandleMpvCommandFromIpcMainDepsHandler
>[0];
type RunSubsyncManualFromIpcMainDeps<TRequest, TResult> = Parameters<
  typeof createBuildRunSubsyncManualFromIpcMainDepsHandler<TRequest, TResult>
>[0];

export function createIpcRuntimeHandlers<TRequest, TResult>(deps: {
  handleMpvCommandFromIpcDeps: HandleMpvCommandFromIpcMainDeps;
  runSubsyncManualFromIpcDeps: RunSubsyncManualFromIpcMainDeps<TRequest, TResult>;
}) {
  const handleMpvCommandFromIpcMainDeps = createBuildHandleMpvCommandFromIpcMainDepsHandler(
    deps.handleMpvCommandFromIpcDeps,
  )();
  const handleMpvCommandFromIpc = createHandleMpvCommandFromIpcHandler(
    handleMpvCommandFromIpcMainDeps,
  );

  const runSubsyncManualFromIpcMainDeps = createBuildRunSubsyncManualFromIpcMainDepsHandler<
    TRequest,
    TResult
  >(deps.runSubsyncManualFromIpcDeps)();
  const runSubsyncManualFromIpc = createRunSubsyncManualFromIpcHandler<TRequest, TResult>(
    runSubsyncManualFromIpcMainDeps,
  );

  return {
    handleMpvCommandFromIpc,
    runSubsyncManualFromIpc,
  };
}
