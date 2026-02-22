import type { RegisterIpcRuntimeServicesParams } from '../../ipc-runtime';
import type { SubsyncManualRunRequest, SubsyncResult } from '../../../types';
import {
  createBuildMpvCommandFromIpcRuntimeMainDepsHandler,
  createIpcRuntimeHandlers,
} from '../domains/ipc';
import type { ComposerInputs, ComposerOutputs } from './contracts';

type MpvCommand = (string | number)[];

type IpcMainDeps = RegisterIpcRuntimeServicesParams['mainDeps'];
type IpcMainDepsWithoutHandlers = Omit<IpcMainDeps, 'handleMpvCommand' | 'runSubsyncManual'>;
type RunSubsyncManual = IpcMainDeps['runSubsyncManual'];

type IpcRuntimeDeps = Parameters<
  typeof createIpcRuntimeHandlers<SubsyncManualRunRequest, SubsyncResult>
>[0];

export type IpcRuntimeComposerOptions = ComposerInputs<{
  mpvCommandMainDeps: Parameters<typeof createBuildMpvCommandFromIpcRuntimeMainDepsHandler>[0];
  handleMpvCommandFromIpcRuntime: IpcRuntimeDeps['handleMpvCommandFromIpcDeps']['handleMpvCommandFromIpcRuntime'];
  runSubsyncManualFromIpc: RunSubsyncManual;
  registration: {
    runtimeOptions: RegisterIpcRuntimeServicesParams['runtimeOptions'];
    mainDeps: IpcMainDepsWithoutHandlers;
    ankiJimakuDeps: RegisterIpcRuntimeServicesParams['ankiJimakuDeps'];
    registerIpcRuntimeServices: (params: RegisterIpcRuntimeServicesParams) => void;
  };
}>;

export type IpcRuntimeComposerResult = ComposerOutputs<{
  handleMpvCommandFromIpc: (command: MpvCommand) => void;
  runSubsyncManualFromIpc: RunSubsyncManual;
  registerIpcRuntimeHandlers: () => void;
}>;

export function composeIpcRuntimeHandlers(
  options: IpcRuntimeComposerOptions,
): IpcRuntimeComposerResult {
  const mpvCommandFromIpcRuntimeMainDeps = createBuildMpvCommandFromIpcRuntimeMainDepsHandler(
    options.mpvCommandMainDeps,
  )();
  const { handleMpvCommandFromIpc, runSubsyncManualFromIpc } = createIpcRuntimeHandlers<
    SubsyncManualRunRequest,
    SubsyncResult
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
        runSubsyncManual: (request) => runSubsyncManualFromIpc(request),
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
