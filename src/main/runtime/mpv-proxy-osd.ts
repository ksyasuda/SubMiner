type MpvPropertyClientLike = {
  connected: boolean;
  requestProperty: (name: string) => Promise<unknown>;
};

type MpvSubtitleTrack = {
  id?: number;
  type?: string;
  selected?: boolean;
  external?: boolean;
  lang?: string;
  title?: string;
};

function getTrackOsdPrefix(command: (string | number)[]): string | null {
  const operation = typeof command[0] === 'string' ? command[0] : '';
  const property = typeof command[1] === 'string' ? command[1] : '';
  const modifiesProperty =
    operation === 'add' ||
    operation === 'set' ||
    operation === 'set_property' ||
    operation === 'cycle' ||
    operation === 'cycle-values' ||
    operation === 'multiply';
  if (!modifiesProperty) return null;
  if (property === 'sid') return 'Subtitle track';
  if (property === 'secondary-sid') return 'Secondary subtitle track';
  return null;
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

function normalizeTrackList(trackListRaw: unknown): MpvSubtitleTrack[] {
  if (!Array.isArray(trackListRaw)) return [];
  return trackListRaw
    .filter(
      (track): track is Record<string, unknown> => Boolean(track) && typeof track === 'object',
    )
    .map((track) => {
      const id = parseTrackId(track.id);
      return {
        ...track,
        id: id === null ? undefined : id,
      } as MpvSubtitleTrack;
    });
}

function formatSubtitleTrackLabel(track: MpvSubtitleTrack): string {
  const trackId = typeof track.id === 'number' ? track.id : -1;
  const source = track.external ? 'External' : 'Internal';
  const label = track.lang || track.title || 'unknown';
  const active = track.selected ? ' (active)' : '';
  return `${source} #${trackId} - ${label}${active}`;
}

export async function resolveProxyCommandOsdRuntime(
  command: (string | number)[],
  getMpvClient: () => MpvPropertyClientLike | null,
): Promise<string | null> {
  const prefix = getTrackOsdPrefix(command);
  if (!prefix) return null;

  const client = getMpvClient();
  if (!client?.connected) return null;

  const property = prefix === 'Subtitle track' ? 'sid' : 'secondary-sid';
  const [trackListRaw, trackIdRaw] = await Promise.all([
    client.requestProperty('track-list'),
    client.requestProperty(property),
  ]);

  const trackId = parseTrackId(trackIdRaw);
  if (trackId === null) {
    return `${prefix}: none`;
  }

  const track = normalizeTrackList(trackListRaw).find(
    (entry) => entry.type === 'sub' && entry.id === trackId,
  );
  if (!track) {
    return `${prefix}: #${trackId}`;
  }

  return `${prefix}: ${formatSubtitleTrackLabel(track)}`;
}
