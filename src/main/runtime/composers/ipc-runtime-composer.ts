import type { RegisterIpcRuntimeServicesParams } from '../../ipc-runtime';
import {
  createBuildMpvCommandFromIpcRuntimeMainDepsHandler,
  createIpcRuntimeHandlers,
} from '../domains/ipc';

type MpvCommand = (string | number)[];

type IpcMainDepsWithoutHandlers = Omit<
  RegisterIpcRuntimeServicesParams['mainDeps'],
  'handleMpvCommand' | 'runSubsyncManual'
>;

type IpcRuntimeDeps<TRequest, TResult> = Parameters<
  typeof createIpcRuntimeHandlers<TRequest, TResult>
>[0];

export type IpcRuntimeComposerOptions<TRequest, TResult> = {
  mpvCommandMainDeps: Parameters<typeof createBuildMpvCommandFromIpcRuntimeMainDepsHandler>[0];
  handleMpvCommandFromIpcRuntime: IpcRuntimeDeps<
    TRequest,
    TResult
  >['handleMpvCommandFromIpcDeps']['handleMpvCommandFromIpcRuntime'];
  runSubsyncManualFromIpc: (request: TRequest) => Promise<TResult>;
  registration: {
    runtimeOptions: RegisterIpcRuntimeServicesParams['runtimeOptions'];
    mainDeps: IpcMainDepsWithoutHandlers;
    ankiJimakuDeps: RegisterIpcRuntimeServicesParams['ankiJimakuDeps'];
    registerIpcRuntimeServices: (params: RegisterIpcRuntimeServicesParams) => void;
  };
};

export type IpcRuntimeComposerResult<TRequest, TResult> = {
  handleMpvCommandFromIpc: (command: MpvCommand) => void;
  runSubsyncManualFromIpc: (request: TRequest) => Promise<TResult>;
  registerIpcRuntimeHandlers: () => void;
};

export function composeIpcRuntimeHandlers<TRequest, TResult>(
  options: IpcRuntimeComposerOptions<TRequest, TResult>,
): IpcRuntimeComposerResult<TRequest, TResult> {
  const mpvCommandFromIpcRuntimeMainDeps = createBuildMpvCommandFromIpcRuntimeMainDepsHandler(
    options.mpvCommandMainDeps,
  )();
  const { handleMpvCommandFromIpc, runSubsyncManualFromIpc } = createIpcRuntimeHandlers<
    TRequest,
    TResult
  >({
    handleMpvCommandFromIpcDeps: {
      handleMpvCommandFromIpcRuntime: options.handleMpvCommandFromIpcRuntime,
      buildMpvCommandDeps: () => mpvCommandFromIpcRuntimeMainDeps,
    },
    runSubsyncManualFromIpcDeps: {
      runManualFromIpc: (request) => options.runSubsyncManualFromIpc(request),
    },
  });

  const registerIpcRuntimeHandlers = (): void => {
    options.registration.registerIpcRuntimeServices({
      runtimeOptions: options.registration.runtimeOptions,
      mainDeps: {
        ...options.registration.mainDeps,
        handleMpvCommand: (command) => handleMpvCommandFromIpc(command),
        runSubsyncManual: (request) => runSubsyncManualFromIpc(request as TRequest),
      },
      ankiJimakuDeps: options.registration.ankiJimakuDeps,
    });
  };

  return {
    handleMpvCommandFromIpc: (command) => handleMpvCommandFromIpc(command),
    runSubsyncManualFromIpc: (request) => runSubsyncManualFromIpc(request),
    registerIpcRuntimeHandlers,
  };
}
