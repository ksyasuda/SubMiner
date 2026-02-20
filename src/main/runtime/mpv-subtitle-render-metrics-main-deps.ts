import type { createUpdateMpvSubtitleRenderMetricsHandler } from './mpv-subtitle-render-metrics';

type UpdateMpvSubtitleRenderMetricsMainDeps = Parameters<typeof createUpdateMpvSubtitleRenderMetricsHandler>[0];

export function createBuildUpdateMpvSubtitleRenderMetricsMainDepsHandler(
  deps: UpdateMpvSubtitleRenderMetricsMainDeps,
) {
  return (): UpdateMpvSubtitleRenderMetricsMainDeps => ({
    getCurrentMetrics: () => deps.getCurrentMetrics(),
    setCurrentMetrics: (metrics) => deps.setCurrentMetrics(metrics),
    applyPatch: (current, patch) => deps.applyPatch(current, patch),
    broadcastMetrics: (metrics) => deps.broadcastMetrics(metrics),
  });
}
