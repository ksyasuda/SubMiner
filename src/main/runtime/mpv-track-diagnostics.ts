type MpvTrackDiagnosticEntry = {
  id: number | null;
  type: string | null;
  selected: boolean;
  external: boolean;
  lang: string | null;
  title: string | null;
  codec: string | null;
};

export type SubtitleTrackDiagnostics = {
  trackListReadable: boolean;
  trackCount: number;
  subtitleTrackCount: number;
  activePrimarySid: number | null;
  selectedSubtitleIds: number[];
  externalSubtitleCount: number;
  internalSubtitleCount: number;
  languages: string[];
  selectedSubtitleLabels: string[];
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object');
}

function parseTrackId(value: unknown): number | null {
  if (typeof value === 'number' && Number.isInteger(value)) {
    return value;
  }
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed.length || trimmed === 'no' || trimmed === 'auto') {
      return null;
    }
    const parsed = Number(trimmed);
    if (Number.isInteger(parsed)) {
      return parsed;
    }
  }
  return null;
}

function readString(value: unknown): string | null {
  return typeof value === 'string' && value.trim().length > 0 ? value.trim() : null;
}

function normalizeTrack(track: unknown): MpvTrackDiagnosticEntry | null {
  if (!isRecord(track)) {
    return null;
  }

  return {
    id: parseTrackId(track.id),
    type: readString(track.type),
    selected: track.selected === true,
    external: track.external === true,
    lang: readString(track.lang),
    title: readString(track.title),
    codec: readString(track.codec),
  };
}

function formatSubtitleTrackLabel(track: MpvTrackDiagnosticEntry): string {
  const id = track.id === null ? '?' : String(track.id);
  const source = track.external ? 'external' : 'internal';
  const label = track.lang ?? track.title ?? track.codec ?? 'unknown';
  return `${source}#${id}:${label}`;
}

export function buildSubtitleTrackDiagnostics(
  activePrimarySid: number | null,
  trackList: unknown[] | null,
): SubtitleTrackDiagnostics {
  if (!Array.isArray(trackList)) {
    return {
      trackListReadable: false,
      trackCount: 0,
      subtitleTrackCount: 0,
      activePrimarySid,
      selectedSubtitleIds: [],
      externalSubtitleCount: 0,
      internalSubtitleCount: 0,
      languages: [],
      selectedSubtitleLabels: [],
    };
  }

  const normalizedTracks = trackList.map(normalizeTrack).filter((track) => track !== null);
  const subtitleTracks = normalizedTracks.filter((track) => track.type === 'sub');
  const selectedSubtitleTracks = subtitleTracks.filter((track) => track.selected);
  const languages = Array.from(
    new Set(
      subtitleTracks
        .map((track) => track.lang)
        .filter((language): language is string => language !== null),
    ),
  ).sort((left, right) => left.localeCompare(right));

  return {
    trackListReadable: true,
    trackCount: normalizedTracks.length,
    subtitleTrackCount: subtitleTracks.length,
    activePrimarySid,
    selectedSubtitleIds: selectedSubtitleTracks
      .map((track) => track.id)
      .filter((id): id is number => id !== null),
    externalSubtitleCount: subtitleTracks.filter((track) => track.external).length,
    internalSubtitleCount: subtitleTracks.filter((track) => !track.external).length,
    languages,
    selectedSubtitleLabels: selectedSubtitleTracks.map(formatSubtitleTrackLabel),
  };
}
