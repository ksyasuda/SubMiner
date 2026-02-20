import { createBuildInitializeOverlayRuntimeBootstrapMainDepsHandler } from './app-runtime-main-deps';
import { createInitializeOverlayRuntimeHandler } from './overlay-runtime-bootstrap';
import { createBuildInitializeOverlayRuntimeOptionsHandler } from './overlay-runtime-options';
import { createBuildInitializeOverlayRuntimeMainDepsHandler } from './overlay-runtime-options-main-deps';

type InitializeOverlayRuntimeMainDeps = Parameters<
  typeof createBuildInitializeOverlayRuntimeMainDepsHandler
>[0];
type InitializeOverlayRuntimeBootstrapMainDeps = Parameters<
  typeof createBuildInitializeOverlayRuntimeBootstrapMainDepsHandler
>[0];

export function createOverlayRuntimeBootstrapHandlers(deps: {
  initializeOverlayRuntimeMainDeps: InitializeOverlayRuntimeMainDeps;
  initializeOverlayRuntimeBootstrapDeps: Omit<InitializeOverlayRuntimeBootstrapMainDeps, 'buildOptions'>;
}) {
  const buildInitializeOverlayRuntimeOptionsHandler = createBuildInitializeOverlayRuntimeOptionsHandler(
    createBuildInitializeOverlayRuntimeMainDepsHandler(deps.initializeOverlayRuntimeMainDeps)(),
  );
  const initializeOverlayRuntime = createInitializeOverlayRuntimeHandler(
    createBuildInitializeOverlayRuntimeBootstrapMainDepsHandler({
      ...deps.initializeOverlayRuntimeBootstrapDeps,
      buildOptions: () => buildInitializeOverlayRuntimeOptionsHandler(),
    })(),
  );

  return {
    initializeOverlayRuntime,
  };
}
