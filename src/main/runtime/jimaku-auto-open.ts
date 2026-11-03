import {
  normalizeLocalMediaPath,
  resolveManagedLocalSubtitleSelection,
} from './local-subtitle-selection';

// mpv lists sidecar and embedded subtitle tracks right after the file loads; the delay lets that
// track list settle before deciding nothing matches.
const DEFAULT_CHECK_DELAY_MS = 1000;

/**
 * Opens the Jimaku modal for local media that has no subtitle track in the primary subtitle
 * languages (Japanese by default). Checks once per media load and pauses playback first, so the
 * episode does not start without subtitles.
 */
export function createJimakuAutoOpenRuntime(deps: {
  isEnabled: () => boolean;
  getCurrentMediaPath: () => string | null;
  getMpvClient: () => {
    connected?: boolean;
    requestProperty?: (name: string) => Promise<unknown>;
  } | null;
  getPrimarySubtitleLanguages: () => string[];
  holdPlayback: () => void;
  openJimaku: () => void;
  schedule: (callback: () => void, delayMs: number) => ReturnType<typeof setTimeout>;
  clearScheduled: (timer: ReturnType<typeof setTimeout>) => void;
  delayMs?: number;
}) {
  const delayMs = deps.delayMs ?? DEFAULT_CHECK_DELAY_MS;
  let currentMediaPath: string | null = null;
  let pendingTimer: ReturnType<typeof setTimeout> | null = null;

  const clearPendingTimer = (): void => {
    if (!pendingTimer) {
      return;
    }
    deps.clearScheduled(pendingTimer);
    pendingTimer = null;
  };

  const checkMediaPath = async (mediaPath: string): Promise<void> => {
    const client = deps.getMpvClient();
    if (!client?.connected || !client.requestProperty) {
      return;
    }
    let trackList: unknown;
    try {
      trackList = await client.requestProperty('track-list');
    } catch {
      // Never open the modal on an unreadable track list; a missing track is not proven.
      return;
    }
    if (
      mediaPath !== currentMediaPath ||
      normalizeLocalMediaPath(deps.getCurrentMediaPath()) !== mediaPath ||
      !deps.isEnabled()
    ) {
      return;
    }
    const { hasPrimaryMatch } = resolveManagedLocalSubtitleSelection({
      trackList: Array.isArray(trackList) ? trackList : null,
      primaryLanguages: deps.getPrimarySubtitleLanguages(),
      secondaryLanguages: [],
    });
    if (hasPrimaryMatch) {
      return;
    }
    deps.holdPlayback();
    deps.openJimaku();
  };

  return {
    handleMediaPathChange: (mediaPath: string | null | undefined): void => {
      const normalizedPath = normalizeLocalMediaPath(mediaPath);
      if (normalizedPath === currentMediaPath) {
        return;
      }
      clearPendingTimer();
      currentMediaPath = normalizedPath;
      if (!normalizedPath || !deps.isEnabled()) {
        return;
      }
      pendingTimer = deps.schedule(() => {
        pendingTimer = null;
        void checkMediaPath(normalizedPath);
      }, delayMs);
    },
  };
}
