import type { MpvSubtitleRenderMetrics } from '../../types';

export function createUpdateMpvSubtitleRenderMetricsHandler(deps: {
  getCurrentMetrics: () => MpvSubtitleRenderMetrics;
  setCurrentMetrics: (metrics: MpvSubtitleRenderMetrics) => void;
  applyPatch: (
    current: MpvSubtitleRenderMetrics,
    patch: Partial<MpvSubtitleRenderMetrics>,
  ) => { next: MpvSubtitleRenderMetrics; changed: boolean };
  broadcastMetrics: (metrics: MpvSubtitleRenderMetrics) => void;
}) {
  return (patch: Partial<MpvSubtitleRenderMetrics>): void => {
    const { next, changed } = deps.applyPatch(deps.getCurrentMetrics(), patch);
    if (!changed) return;
    deps.setCurrentMetrics(next);
    deps.broadcastMetrics(next);
  };
}
