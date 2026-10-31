import type { CycleSecondarySubModeDeps } from '../../core/services/subtitle-position';

export function createCycleSecondarySubModeRuntimeHandler(deps: {
  cycleSecondarySubModeMainDeps: CycleSecondarySubModeDeps;
  cycleSecondarySubMode: (deps: CycleSecondarySubModeDeps) => void;
}) {
  return () => deps.cycleSecondarySubMode(deps.cycleSecondarySubModeMainDeps);
}
