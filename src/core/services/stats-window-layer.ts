export type StatsWindowLayerSuspensionState = {
  count: number;
};

export function createStatsWindowLayerSuspensionState(): StatsWindowLayerSuspensionState {
  return { count: 0 };
}

export function isStatsWindowLayerSuspended(state: StatsWindowLayerSuspensionState): boolean {
  return state.count > 0;
}

export function suspendStatsWindowLayer(state: StatsWindowLayerSuspensionState): boolean {
  state.count += 1;
  return state.count === 1;
}

export function restoreStatsWindowLayer(state: StatsWindowLayerSuspensionState): boolean {
  if (state.count <= 0) {
    return false;
  }

  state.count -= 1;
  return state.count === 0;
}

export function resetStatsWindowLayerSuspension(state: StatsWindowLayerSuspensionState): void {
  state.count = 0;
}
