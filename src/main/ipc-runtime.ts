import {
  createIpcDepsRuntimeService,
  registerAnkiJimakuIpcRuntimeService,
  registerIpcHandlersService,
} from "../core/services";
import { registerAnkiJimakuIpcHandlers } from "../core/services/anki-jimaku-ipc-service";
import {
  createAnkiJimakuIpcRuntimeServiceDeps,
  AnkiJimakuIpcRuntimeServiceDepsParams,
  createMainIpcRuntimeServiceDeps,
  MainIpcRuntimeServiceDepsParams,
  createRuntimeOptionsIpcDeps,
  RuntimeOptionsIpcDepsParams,
} from "./dependencies";

export interface RegisterIpcRuntimeServicesParams {
  runtimeOptions: RuntimeOptionsIpcDepsParams;
  mainDeps: Omit<
    MainIpcRuntimeServiceDepsParams,
    "setRuntimeOption" | "cycleRuntimeOption"
  >;
  ankiJimakuDeps: AnkiJimakuIpcRuntimeServiceDepsParams;
}

export function registerMainIpcRuntimeServices(
  params: MainIpcRuntimeServiceDepsParams,
): void {
  registerIpcHandlersService(
    createIpcDepsRuntimeService(createMainIpcRuntimeServiceDeps(params)),
  );
}

export function registerAnkiJimakuIpcRuntimeServices(
  params: AnkiJimakuIpcRuntimeServiceDepsParams,
): void {
  registerAnkiJimakuIpcRuntimeService(
    createAnkiJimakuIpcRuntimeServiceDeps(params),
    registerAnkiJimakuIpcHandlers,
  );
}

export function registerIpcRuntimeServices(
  params: RegisterIpcRuntimeServicesParams,
): void {
  const runtimeOptionsIpcDeps = createRuntimeOptionsIpcDeps({
    getRuntimeOptionsManager: params.runtimeOptions.getRuntimeOptionsManager,
    showMpvOsd: params.runtimeOptions.showMpvOsd,
  });
  registerMainIpcRuntimeServices({
    ...params.mainDeps,
    setRuntimeOption: runtimeOptionsIpcDeps.setRuntimeOption,
    cycleRuntimeOption: runtimeOptionsIpcDeps.cycleRuntimeOption,
  });
  registerAnkiJimakuIpcRuntimeServices(params.ankiJimakuDeps);
}
