import { createHandleInitialArgsHandler } from './initial-args-handler';

type InitialArgsMainDeps = Parameters<typeof createHandleInitialArgsHandler>[0];

export function createInitialArgsRuntimeHandler(deps: InitialArgsMainDeps) {
  return createHandleInitialArgsHandler(deps);
}
