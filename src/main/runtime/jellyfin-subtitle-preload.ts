type JellyfinSession = {
  serverUrl: string;
  accessToken: string;
  userId: string;
  username: string;
};

type JellyfinClientInfo = {
  clientName: string;
  clientVersion: string;
  deviceId: string;
};

type JellyfinSubtitleTrack = {
  index: number;
  language?: string;
  title?: string;
  deliveryUrl?: string | null;
};

type MpvClientLike = {
  requestProperty: (name: string) => Promise<unknown>;
};

function normalizeLang(value: unknown): string {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/_/g, '-');
}

function isJapanese(value: string): boolean {
  const v = normalizeLang(value);
  return (
    v === 'ja' ||
    v === 'jp' ||
    v === 'jpn' ||
    v === 'japanese' ||
    v.startsWith('ja-') ||
    v.startsWith('jp-')
  );
}

function isEnglish(value: string): boolean {
  const v = normalizeLang(value);
  return (
    v === 'en' ||
    v === 'eng' ||
    v === 'english' ||
    v === 'enus' ||
    v === 'en-us' ||
    v.startsWith('en-')
  );
}

function isLikelyHearingImpaired(title: string): boolean {
  return /\b(hearing impaired|sdh|closed captions?|cc)\b/i.test(title);
}

function pickBestTrackId(
  tracks: Array<{
    id: number;
    lang: string;
    title: string;
    external: boolean;
  }>,
  languageMatcher: (value: string) => boolean,
  excludeId: number | null = null,
): number | null {
  const ranked = tracks
    .filter((track) => languageMatcher(track.lang))
    .filter((track) => track.id !== excludeId)
    .map((track) => ({
      track,
      score:
        (track.external ? 100 : 0) +
        (isLikelyHearingImpaired(track.title) ? -10 : 10) +
        (/\bdefault\b/i.test(track.title) ? 3 : 0),
    }))
    .sort((a, b) => b.score - a.score);
  return ranked[0]?.track.id ?? null;
}

export function createPreloadJellyfinExternalSubtitlesHandler(deps: {
  listJellyfinSubtitleTracks: (
    session: JellyfinSession,
    clientInfo: JellyfinClientInfo,
    itemId: string,
  ) => Promise<JellyfinSubtitleTrack[]>;
  getMpvClient: () => MpvClientLike | null;
  sendMpvCommand: (command: Array<string | number>) => void;
  wait: (ms: number) => Promise<void>;
  logDebug: (message: string, error: unknown) => void;
}) {
  return async (params: {
    session: JellyfinSession;
    clientInfo: JellyfinClientInfo;
    itemId: string;
  }): Promise<void> => {
    try {
      const tracks = await deps.listJellyfinSubtitleTracks(
        params.session,
        params.clientInfo,
        params.itemId,
      );
      const externalTracks = tracks.filter((track) => Boolean(track.deliveryUrl));
      if (externalTracks.length === 0) {
        return;
      }

      await deps.wait(300);
      const seenUrls = new Set<string>();
      for (const track of externalTracks) {
        if (!track.deliveryUrl || seenUrls.has(track.deliveryUrl)) {
          continue;
        }
        seenUrls.add(track.deliveryUrl);
        const labelBase = (track.title || track.language || '').trim();
        const label = labelBase || `Jellyfin Subtitle ${track.index}`;
        deps.sendMpvCommand(['sub-add', track.deliveryUrl, 'cached', label, track.language || '']);
      }

      await deps.wait(250);
      const trackListRaw = await deps.getMpvClient()?.requestProperty('track-list');
      const subtitleTracks = Array.isArray(trackListRaw)
        ? trackListRaw
            .filter(
              (track): track is Record<string, unknown> =>
                Boolean(track) &&
                typeof track === 'object' &&
                track.type === 'sub' &&
                typeof track.id === 'number',
            )
            .map((track) => ({
              id: track.id as number,
              lang: String(track.lang || ''),
              title: String(track.title || ''),
              external: track.external === true,
            }))
        : [];

      const japanesePrimaryId = pickBestTrackId(subtitleTracks, isJapanese);
      if (japanesePrimaryId !== null) {
        deps.sendMpvCommand(['set_property', 'sid', japanesePrimaryId]);
      } else {
        deps.sendMpvCommand(['set_property', 'sid', 'no']);
      }

      const englishSecondaryId = pickBestTrackId(subtitleTracks, isEnglish, japanesePrimaryId);
      if (englishSecondaryId !== null) {
        deps.sendMpvCommand(['set_property', 'secondary-sid', englishSecondaryId]);
      }
    } catch (error) {
      deps.logDebug('Failed to preload Jellyfin external subtitles', error);
    }
  };
}
