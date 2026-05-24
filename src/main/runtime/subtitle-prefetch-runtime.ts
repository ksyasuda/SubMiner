import type { SubtitlePrefetchInitController } from './subtitle-prefetch-init';
import { buildSubtitleSidebarSourceKey } from './subtitle-prefetch-source';

type MpvSubtitleTrackLike = {
  type?: unknown;
  id?: unknown;
  selected?: unknown;
  external?: unknown;
  codec?: unknown;
  'ff-index'?: unknown;
  'external-filename'?: unknown;
};

type ActiveSubtitleSidebarSource = {
  path: string;
  sourceKey: string;
  cleanup?: () => Promise<void>;
};

function parseTrackId(value: unknown): number | null {
  if (typeof value === 'number' && Number.isInteger(value)) {
    return value;
  }
  if (typeof value === 'string') {
    const parsed = Number(value.trim());
    return Number.isInteger(parsed) ? parsed : null;
  }
  return null;
}

function getActiveSubtitleTrack(
  currentTrackRaw: unknown,
  trackListRaw: unknown,
  sidRaw: unknown,
): MpvSubtitleTrackLike | null {
  if (currentTrackRaw && typeof currentTrackRaw === 'object') {
    const track = currentTrackRaw as MpvSubtitleTrackLike;
    if (track.type === undefined || track.type === 'sub') {
      return track;
    }
  }

  const sid = parseTrackId(sidRaw);
  if (!Array.isArray(trackListRaw)) {
    return null;
  }

  const bySid =
    sid === null
      ? null
      : ((trackListRaw.find((entry: unknown) => {
          if (!entry || typeof entry !== 'object') {
            return false;
          }
          const track = entry as MpvSubtitleTrackLike;
          return track.type === 'sub' && parseTrackId(track.id) === sid;
        }) as MpvSubtitleTrackLike | undefined) ?? null);
  if (bySid) {
    return bySid;
  }

  return (
    (trackListRaw.find((entry: unknown) => {
      if (!entry || typeof entry !== 'object') {
        return false;
      }
      const track = entry as MpvSubtitleTrackLike;
      return track.type === 'sub' && track.selected === true;
    }) as MpvSubtitleTrackLike | undefined) ?? null
  );
}

export function createResolveActiveSubtitleSidebarSourceHandler(deps: {
  getFfmpegPath: () => string;
  extractInternalSubtitleTrack: (
    ffmpegPath: string,
    videoPath: string,
    track: MpvSubtitleTrackLike,
  ) => Promise<{ path: string; cleanup: () => Promise<void> } | null>;
}) {
  return async (input: {
    currentExternalFilenameRaw: unknown;
    currentTrackRaw: unknown;
    trackListRaw: unknown;
    sidRaw: unknown;
    videoPath: string;
  }): Promise<ActiveSubtitleSidebarSource | null> => {
    const currentExternalFilename =
      typeof input.currentExternalFilenameRaw === 'string'
        ? input.currentExternalFilenameRaw.trim()
        : '';
    if (currentExternalFilename) {
      return { path: currentExternalFilename, sourceKey: currentExternalFilename };
    }

    const track = getActiveSubtitleTrack(input.currentTrackRaw, input.trackListRaw, input.sidRaw);
    if (!track) {
      return null;
    }

    const externalFilename =
      typeof track['external-filename'] === 'string' ? track['external-filename'].trim() : '';
    if (externalFilename) {
      return { path: externalFilename, sourceKey: externalFilename };
    }

    const extracted = await deps.extractInternalSubtitleTrack(
      deps.getFfmpegPath(),
      input.videoPath,
      track,
    );
    if (!extracted) {
      return null;
    }

    return {
      ...extracted,
      sourceKey: buildSubtitleSidebarSourceKey(input.videoPath, track, extracted.path),
    };
  };
}

export function createRefreshSubtitlePrefetchFromActiveTrackHandler(deps: {
  getMpvClient: () => {
    connected?: boolean;
    requestProperty: (name: string) => Promise<unknown>;
  } | null;
  getLastObservedTimePos: () => number;
  shouldKeepExistingCuesOnMissingSource?: (videoPath: string) => boolean;
  subtitlePrefetchInitController: SubtitlePrefetchInitController;
  resolveActiveSubtitleSidebarSource: (
    input: Parameters<ReturnType<typeof createResolveActiveSubtitleSidebarSourceHandler>>[0],
  ) => Promise<ActiveSubtitleSidebarSource | null>;
}) {
  return async (): Promise<void> => {
    const client = deps.getMpvClient();
    if (!client?.connected) {
      return;
    }

    try {
      const [currentExternalFilenameRaw, currentTrackRaw, trackListRaw, sidRaw, videoPathRaw] =
        await Promise.all([
          client.requestProperty('current-tracks/sub/external-filename').catch(() => null),
          client.requestProperty('current-tracks/sub').catch(() => null),
          client.requestProperty('track-list'),
          client.requestProperty('sid'),
          client.requestProperty('path'),
        ]);
      const videoPath = typeof videoPathRaw === 'string' ? videoPathRaw : '';
      if (!videoPath) {
        deps.subtitlePrefetchInitController.cancelPendingInit();
        return;
      }

      const resolvedSource = await deps.resolveActiveSubtitleSidebarSource({
        currentExternalFilenameRaw,
        currentTrackRaw,
        trackListRaw,
        sidRaw,
        videoPath,
      });
      if (!resolvedSource) {
        if (deps.shouldKeepExistingCuesOnMissingSource?.(videoPath) === true) {
          return;
        }
        deps.subtitlePrefetchInitController.cancelPendingInit();
        return;
      }

      try {
        await deps.subtitlePrefetchInitController.initSubtitlePrefetch(
          resolvedSource.path,
          deps.getLastObservedTimePos(),
          resolvedSource.sourceKey,
        );
      } finally {
        await resolvedSource.cleanup?.();
      }
    } catch {
      // Skip refresh when the track query fails.
    }
  };
}
