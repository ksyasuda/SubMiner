import {
  RuntimeOptionId,
  RuntimeOptionValue,
} from "../../types";
import {
  cycleRuntimeOptionFromIpcRuntimeService,
  RuntimeOptionsManagerLike,
  setRuntimeOptionFromIpcRuntimeService,
} from "./runtime-options-runtime-service";

export interface RuntimeOptionsIpcDepsRuntimeOptions {
  getRuntimeOptionsManager: () => RuntimeOptionsManagerLike | null;
  showMpvOsd: (text: string) => void;
}

export function createRuntimeOptionsIpcDepsRuntimeService(
  options: RuntimeOptionsIpcDepsRuntimeOptions,
): {
  setRuntimeOption: (id: string, value: unknown) => unknown;
  cycleRuntimeOption: (id: string, direction: 1 | -1) => unknown;
} {
  return {
    setRuntimeOption: (id, value) =>
      setRuntimeOptionFromIpcRuntimeService(
        options.getRuntimeOptionsManager(),
        id as RuntimeOptionId,
        value as RuntimeOptionValue,
        options.showMpvOsd,
      ),
    cycleRuntimeOption: (id, direction) =>
      cycleRuntimeOptionFromIpcRuntimeService(
        options.getRuntimeOptionsManager(),
        id as RuntimeOptionId,
        direction,
        options.showMpvOsd,
      ),
  };
}
