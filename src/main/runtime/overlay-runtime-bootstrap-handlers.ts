import { createInitializeOverlayRuntimeHandler } from './overlay-runtime-bootstrap';
import { createBuildInitializeOverlayRuntimeOptionsHandler } from './overlay-runtime-options';

type InitializeOverlayRuntimeDeps = Parameters<typeof createInitializeOverlayRuntimeHandler>[0];

export function createOverlayRuntimeBootstrapHandlers(deps: {
  initializeOverlayRuntimeMainDeps: Parameters<
    typeof createBuildInitializeOverlayRuntimeOptionsHandler
  >[0];
  initializeOverlayRuntimeBootstrapDeps: Omit<InitializeOverlayRuntimeDeps, 'buildOptions'>;
}) {
  const buildInitializeOverlayRuntimeOptions = createBuildInitializeOverlayRuntimeOptionsHandler(
    deps.initializeOverlayRuntimeMainDeps,
  );
  const initializeOverlayRuntime = createInitializeOverlayRuntimeHandler({
    ...deps.initializeOverlayRuntimeBootstrapDeps,
    buildOptions: () => buildInitializeOverlayRuntimeOptions(),
  });

  return {
    initializeOverlayRuntime,
  };
}
