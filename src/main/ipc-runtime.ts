import {
  createIpcDepsRuntimeService,
  registerAnkiJimakuIpcRuntimeService,
  registerIpcHandlersService,
} from "../core/services";
import {
  createAnkiJimakuIpcRuntimeServiceDeps,
  AnkiJimakuIpcRuntimeServiceDepsParams,
  createMainIpcRuntimeServiceDeps,
  MainIpcRuntimeServiceDepsParams,
} from "./dependencies";

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
  );
}

