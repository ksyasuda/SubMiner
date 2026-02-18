import { updateCurrentMediaPath } from '../core/services';

import type { SubtitlePosition } from '../types';

export interface MediaRuntimeDeps {
  isRemoteMediaPath: (mediaPath: string) => boolean;
  loadSubtitlePosition: () => SubtitlePosition | null;
  getCurrentMediaPath: () => string | null;
  getPendingSubtitlePosition: () => SubtitlePosition | null;
  getSubtitlePositionsDir: () => string;
  setCurrentMediaPath: (mediaPath: string | null) => void;
  clearPendingSubtitlePosition: () => void;
  setSubtitlePosition: (position: SubtitlePosition | null) => void;
  broadcastSubtitlePosition: (position: SubtitlePosition | null) => void;
  getCurrentMediaTitle: () => string | null;
  setCurrentMediaTitle: (title: string | null) => void;
}

export interface MediaRuntimeService {
  updateCurrentMediaPath: (mediaPath: unknown) => void;
  updateCurrentMediaTitle: (mediaTitle: unknown) => void;
  resolveMediaPathForJimaku: (mediaPath: string | null) => string | null;
}

export function createMediaRuntimeService(deps: MediaRuntimeDeps): MediaRuntimeService {
  return {
    updateCurrentMediaPath(mediaPath: unknown): void {
      if (typeof mediaPath !== 'string' || !deps.isRemoteMediaPath(mediaPath)) {
        deps.setCurrentMediaTitle(null);
      }

      updateCurrentMediaPath({
        mediaPath,
        currentMediaPath: deps.getCurrentMediaPath(),
        pendingSubtitlePosition: deps.getPendingSubtitlePosition(),
        subtitlePositionsDir: deps.getSubtitlePositionsDir(),
        loadSubtitlePosition: () => deps.loadSubtitlePosition(),
        setCurrentMediaPath: (nextPath: string | null) => {
          deps.setCurrentMediaPath(nextPath);
        },
        clearPendingSubtitlePosition: () => {
          deps.clearPendingSubtitlePosition();
        },
        setSubtitlePosition: (position: SubtitlePosition | null) => {
          deps.setSubtitlePosition(position);
        },
        broadcastSubtitlePosition: (position: SubtitlePosition | null) => {
          deps.broadcastSubtitlePosition(position);
        },
      });
    },

    updateCurrentMediaTitle(mediaTitle: unknown): void {
      if (typeof mediaTitle === 'string') {
        const sanitized = mediaTitle.trim();
        deps.setCurrentMediaTitle(sanitized.length > 0 ? sanitized : null);
        return;
      }
      deps.setCurrentMediaTitle(null);
    },

    resolveMediaPathForJimaku(mediaPath: string | null): string | null {
      return mediaPath && deps.isRemoteMediaPath(mediaPath) && deps.getCurrentMediaTitle()
        ? deps.getCurrentMediaTitle()
        : mediaPath;
    },
  };
}
