import { createHandleInitialArgsHandler } from './initial-args-handler';
import { createBuildHandleInitialArgsMainDepsHandler } from './initial-args-main-deps';

type InitialArgsMainDeps = Parameters<typeof createBuildHandleInitialArgsMainDepsHandler>[0];

export function createInitialArgsRuntimeHandler(deps: InitialArgsMainDeps) {
  const buildHandleInitialArgsMainDepsHandler = createBuildHandleInitialArgsMainDepsHandler(deps);
  const handleInitialArgsMainDeps = buildHandleInitialArgsMainDepsHandler();
  return createHandleInitialArgsHandler(handleInitialArgsMainDeps);
}
