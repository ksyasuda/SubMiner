import type { SubtitlePosition } from '../../types';

export function createBuildMediaRuntimeMainDepsHandler(deps: {
  isRemoteMediaPath: (mediaPath: string) => boolean;
  loadSubtitlePosition: () => SubtitlePosition | null;
  getCurrentMediaPath: () => string | null;
  getPendingSubtitlePosition: () => SubtitlePosition | null;
  getSubtitlePositionsDir: () => string;
  setCurrentMediaPath: (mediaPath: string | null) => void;
  clearPendingSubtitlePosition: () => void;
  setSubtitlePosition: (position: SubtitlePosition | null) => void;
  broadcastToOverlayWindows: (channel: string, payload: unknown) => void;
  getCurrentMediaTitle: () => string | null;
  setCurrentMediaTitle: (title: string | null) => void;
}) {
  return () => ({
    isRemoteMediaPath: (mediaPath: string) => deps.isRemoteMediaPath(mediaPath),
    loadSubtitlePosition: () => deps.loadSubtitlePosition(),
    getCurrentMediaPath: () => deps.getCurrentMediaPath(),
    getPendingSubtitlePosition: () => deps.getPendingSubtitlePosition(),
    getSubtitlePositionsDir: () => deps.getSubtitlePositionsDir(),
    setCurrentMediaPath: (nextPath: string | null) => deps.setCurrentMediaPath(nextPath),
    clearPendingSubtitlePosition: () => deps.clearPendingSubtitlePosition(),
    setSubtitlePosition: (position: SubtitlePosition | null) => deps.setSubtitlePosition(position),
    broadcastSubtitlePosition: (position: SubtitlePosition | null) =>
      deps.broadcastToOverlayWindows('subtitle-position:set', position),
    getCurrentMediaTitle: () => deps.getCurrentMediaTitle(),
    setCurrentMediaTitle: (title: string | null) => deps.setCurrentMediaTitle(title),
  });
}
