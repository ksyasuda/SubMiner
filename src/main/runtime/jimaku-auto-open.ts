import {
  normalizeLocalMediaPath,
  resolveManagedLocalSubtitleSelection,
} from './local-subtitle-selection';

// mpv lists sidecar and embedded subtitle tracks right after the file loads; the delay lets that
// track list settle before deciding nothing matches.
const DEFAULT_CHECK_DELAY_MS = 1000;
// How long after a failed open a late modal ack still counts as the auto-open showing up.
const LATE_OPEN_ACK_WINDOW_MS = 10_000;

// Audio files report cover art as a video track; only real video means there is something to watch.
function hasVideoTrack(trackList: unknown[]): boolean {
  return trackList.some((entry) => {
    if (!entry || typeof entry !== 'object') {
      return false;
    }
    const track = entry as { type?: unknown; albumart?: unknown; image?: unknown };
    return track.type === 'video' && track.albumart !== true && track.image !== true;
  });
}

type PlaybackHold = {
  pausedByHold: boolean;
  // Set once the renderer acknowledges this hold's open. The renderer reports closes and opens in
  // order, so a close arriving before that acknowledgement belongs to an earlier modal session.
  acknowledged: boolean;
};

/**
 * Opens the Jimaku modal for local media that has no subtitle track in the primary subtitle
 * languages (Japanese by default), checking once per media load.
 *
 * While the auto-opened modal is up, playback is held: paused, and autoplay readiness is deferred
 * instead of resuming it (`isHoldingPlayback`). Closing the modal, or failing to open it, ends the
 * hold: deferred readiness is released and a pause this runtime added is undone, so playback
 * resumes whether or not a subtitle was picked, but a pause made by someone else is kept. A modal
 * that acknowledges shortly after a failed open re-applies the hold.
 */
export function createJimakuAutoOpenRuntime(deps: {
  isEnabled: () => boolean;
  getCurrentMediaPath: () => string | null;
  getMpvClient: () => {
    connected?: boolean;
    requestProperty?: (name: string) => Promise<unknown>;
  } | null;
  getPrimarySubtitleLanguages: () => string[];
  isPlaybackPaused: () => boolean | null;
  pausePlayback: () => void;
  /** Releases readiness deferred during the hold, and unpauses when `undoPause` is set. */
  endHold: (options: { undoPause: boolean }) => void;
  /** Keeps the mpv plugin's pause-until-ready timeout from resuming playback. */
  holdPluginAutoplayGate: () => void;
  /** Resolves to whether the modal acknowledged opening. */
  openJimaku: () => Promise<boolean>;
  schedule: (callback: () => void, delayMs: number) => ReturnType<typeof setTimeout>;
  clearScheduled: (timer: ReturnType<typeof setTimeout>) => void;
  delayMs?: number;
}) {
  const delayMs = deps.delayMs ?? DEFAULT_CHECK_DELAY_MS;
  let currentMediaPath: string | null = null;
  let pendingTimer: ReturnType<typeof setTimeout> | null = null;
  // Media whose auto-opened modal has not closed yet; a late modal ack only counts for it.
  let autoOpenedMediaPath: string | null = null;
  // A fresh object per hold, so a late open failure only ends the hold it started.
  let activeHold: PlaybackHold | null = null;
  let lateAckTimer: ReturnType<typeof setTimeout> | null = null;

  const clearPendingTimer = (): void => {
    if (!pendingTimer) {
      return;
    }
    deps.clearScheduled(pendingTimer);
    pendingTimer = null;
  };

  const clearLateAckTimer = (): void => {
    if (!lateAckTimer) {
      return;
    }
    deps.clearScheduled(lateAckTimer);
    lateAckTimer = null;
  };

  const endHold = (hold: PlaybackHold): void => {
    activeHold = null;
    deps.endHold({ undoPause: hold.pausedByHold });
  };

  const startHold = (): PlaybackHold => {
    const hold = { pausedByHold: deps.isPlaybackPaused() !== true, acknowledged: false };
    activeHold = hold;
    if (hold.pausedByHold) {
      deps.pausePlayback();
    }
    return hold;
  };

  const openForMedia = async (mediaPath: string): Promise<void> => {
    autoOpenedMediaPath = mediaPath;
    const hold = startHold();
    const opened = await deps.openJimaku().catch(() => false);
    if (activeHold !== hold) {
      return;
    }
    if (opened) {
      hold.acknowledged = true;
      // Only once the modal is up: if opening fails, the plugin's own timeout must still resume.
      deps.holdPluginAutoplayGate();
      return;
    }
    endHold(hold);
    // A modal that loads slowly can still show up; past this window an open is the user's own.
    lateAckTimer = deps.schedule(() => {
      lateAckTimer = null;
      if (!activeHold) {
        autoOpenedMediaPath = null;
      }
    }, LATE_OPEN_ACK_WINDOW_MS);
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
      !deps.isEnabled() ||
      !Array.isArray(trackList) ||
      !hasVideoTrack(trackList)
    ) {
      return;
    }
    const { hasPrimaryMatch } = resolveManagedLocalSubtitleSelection({
      trackList,
      primaryLanguages: deps.getPrimarySubtitleLanguages(),
      secondaryLanguages: [],
    });
    if (!hasPrimaryMatch) {
      await openForMedia(mediaPath);
    }
  };

  return {
    handleMediaPathChange: (mediaPath: string | null | undefined): void => {
      const normalizedPath = normalizeLocalMediaPath(mediaPath);
      if (normalizedPath === currentMediaPath) {
        return;
      }
      clearPendingTimer();
      clearLateAckTimer();
      // The gate drops readiness for the old media on its own; the new media is not paused by us.
      activeHold = null;
      autoOpenedMediaPath = null;
      currentMediaPath = normalizedPath;
      if (!normalizedPath || !deps.isEnabled()) {
        return;
      }
      pendingTimer = deps.schedule(() => {
        pendingTimer = null;
        void checkMediaPath(normalizedPath);
      }, delayMs);
    },
    /** The renderer acknowledged the Jimaku modal; re-holds when a timed-out auto-open shows up late. */
    handleJimakuModalOpened: (): void => {
      if (activeHold) {
        activeHold.acknowledged = true;
        return;
      }
      if (!autoOpenedMediaPath || autoOpenedMediaPath !== currentMediaPath) {
        return;
      }
      clearLateAckTimer();
      startHold().acknowledged = true;
      deps.holdPluginAutoplayGate();
    },
    /** The Jimaku modal closed, or the window hosting it went away. */
    handleJimakuModalClosed: (): void => {
      if (activeHold && !activeHold.acknowledged) {
        // A previous modal's close (e.g. the last episode's, reported just as this one opens).
        // If this open never lands, its failure path ends the hold.
        return;
      }
      clearLateAckTimer();
      autoOpenedMediaPath = null;
      if (activeHold) {
        endHold(activeHold);
      }
    },
    isHoldingPlayback: (): boolean => activeHold !== null,
  };
}
