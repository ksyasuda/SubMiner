import { createBuildCycleSecondarySubModeMainDepsHandler } from './secondary-sub-mode-main-deps';

type CycleSecondarySubModeMainDeps = Parameters<
  typeof createBuildCycleSecondarySubModeMainDepsHandler
>[0];
type CycleSecondarySubModeDeps = ReturnType<
  ReturnType<typeof createBuildCycleSecondarySubModeMainDepsHandler>
>;

export function createCycleSecondarySubModeRuntimeHandler(deps: {
  cycleSecondarySubModeMainDeps: CycleSecondarySubModeMainDeps;
  cycleSecondarySubMode: (deps: CycleSecondarySubModeDeps) => void;
}) {
  const buildCycleSecondarySubModeMainDepsHandler = createBuildCycleSecondarySubModeMainDepsHandler(
    deps.cycleSecondarySubModeMainDeps,
  );
  return () => deps.cycleSecondarySubMode(buildCycleSecondarySubModeMainDepsHandler());
}
