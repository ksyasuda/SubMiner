type SubtitleDelayShiftDirection = 'next' | 'previous';

type MpvClientLike = {
  connected: boolean;
  requestProperty: (name: string) => Promise<unknown>;
};

type MpvSubtitleTrackLike = {
  type?: unknown;
  id?: unknown;
  external?: unknown;
  'external-filename'?: unknown;
};

type SubtitleCueCacheEntry = {
  starts: number[];
};

type SubtitleDelayShiftDeps = {
  getMpvClient: () => MpvClientLike | null;
  loadSubtitleSourceText: (source: string) => Promise<string>;
  sendMpvCommand: (command: Array<string | number>) => void;
  showMpvOsd: (text: string) => void;
};

function asTrackId(value: unknown): number | null {
  if (typeof value === 'number' && Number.isInteger(value)) return value;
  if (typeof value === 'string') {
    const parsed = Number(value.trim());
    if (Number.isInteger(parsed)) return parsed;
  }
  return null;
}

function parseSrtOrVttStartTimes(content: string): number[] {
  const starts: number[] = [];
  const lines = content.split(/\r?\n/);
  for (const line of lines) {
    const match = line.match(
      /^\s*(?:(\d{1,2}):)?(\d{2}):(\d{2})[,.](\d{1,3})\s*-->\s*(?:(\d{1,2}):)?(\d{2}):(\d{2})[,.](\d{1,3})/,
    );
    if (!match) continue;
    const hours = Number(match[1] || 0);
    const minutes = Number(match[2] || 0);
    const seconds = Number(match[3] || 0);
    const millis = Number(String(match[4]).padEnd(3, '0'));
    starts.push(hours * 3600 + minutes * 60 + seconds + millis / 1000);
  }
  return starts;
}

function parseAssStartTimes(content: string): number[] {
  const starts: number[] = [];
  const lines = content.split(/\r?\n/);
  for (const line of lines) {
    const match = line.match(/^Dialogue:[^,]*,(\d+:\d{2}:\d{2}\.\d{1,2}),\d+:\d{2}:\d{2}\.\d{1,2},/);
    if (!match) continue;
    const [hoursRaw, minutesRaw, secondsRaw] = match[1]!.split(':');
    if (secondsRaw === undefined) continue;
    const [wholeSecondsRaw, fractionRaw = '0'] = secondsRaw.split('.');
    const hours = Number(hoursRaw);
    const minutes = Number(minutesRaw);
    const wholeSeconds = Number(wholeSecondsRaw);
    const fraction = Number(`0.${fractionRaw}`);
    starts.push(hours * 3600 + minutes * 60 + wholeSeconds + fraction);
  }
  return starts;
}

function normalizeCueStarts(starts: number[]): number[] {
  const sorted = starts
    .filter((value) => Number.isFinite(value) && value >= 0)
    .sort((a, b) => a - b);
  if (sorted.length === 0) return [];

  const deduped: number[] = [sorted[0]!];
  for (let i = 1; i < sorted.length; i += 1) {
    const current = sorted[i]!;
    const previous = deduped[deduped.length - 1]!;
    if (Math.abs(current - previous) > 0.0005) {
      deduped.push(current);
    }
  }
  return deduped;
}

function parseCueStarts(content: string, source: string): number[] {
  const normalizedSource = source.toLowerCase().split('?')[0] || '';
  const parseSrtLike = () => parseSrtOrVttStartTimes(content);
  const parseAssLike = () => parseAssStartTimes(content);

  let starts: number[] = [];
  if (normalizedSource.endsWith('.ass') || normalizedSource.endsWith('.ssa')) {
    starts = parseAssLike();
    if (starts.length === 0) {
      starts = parseSrtLike();
    }
  } else {
    starts = parseSrtLike();
    if (starts.length === 0) {
      starts = parseAssLike();
    }
  }

  const normalized = normalizeCueStarts(starts);
  if (normalized.length === 0) {
    throw new Error('Could not parse subtitle cue timings from active subtitle source.');
  }
  return normalized;
}

function getActiveSubtitleSource(trackListRaw: unknown, sidRaw: unknown): string {
  const sid = asTrackId(sidRaw);
  if (sid === null) {
    throw new Error('No active subtitle track selected.');
  }
  if (!Array.isArray(trackListRaw)) {
    throw new Error('Could not inspect subtitle track list.');
  }

  const activeTrack = trackListRaw.find((entry): entry is MpvSubtitleTrackLike => {
    if (!entry || typeof entry !== 'object') return false;
    const track = entry as MpvSubtitleTrackLike;
    return track.type === 'sub' && asTrackId(track.id) === sid;
  });

  if (!activeTrack) {
    throw new Error('No active subtitle track found in mpv track list.');
  }
  if (activeTrack.external !== true) {
    throw new Error('Active subtitle track is internal and has no direct subtitle file source.');
  }

  const source =
    typeof activeTrack['external-filename'] === 'string'
      ? activeTrack['external-filename'].trim()
      : '';
  if (!source) {
    throw new Error('Active subtitle track has no external subtitle source path.');
  }
  return source;
}

function findAdjacentCueStart(
  starts: number[],
  currentStart: number,
  direction: SubtitleDelayShiftDirection,
): number {
  const epsilon = 0.0005;
  if (direction === 'next') {
    const target = starts.find((value) => value > currentStart + epsilon);
    if (target === undefined) {
      throw new Error('No next subtitle cue found for active subtitle source.');
    }
    return target;
  }

  for (let index = starts.length - 1; index >= 0; index -= 1) {
    const value = starts[index]!;
    if (value < currentStart - epsilon) {
      return value;
    }
  }
  throw new Error('No previous subtitle cue found for active subtitle source.');
}

export function createShiftSubtitleDelayToAdjacentCueHandler(deps: SubtitleDelayShiftDeps) {
  const cueCache = new Map<string, SubtitleCueCacheEntry>();

  return async (direction: SubtitleDelayShiftDirection): Promise<void> => {
    const client = deps.getMpvClient();
    if (!client || !client.connected) {
      throw new Error('MPV not connected.');
    }

    const [trackListRaw, sidRaw, subStartRaw] = await Promise.all([
      client.requestProperty('track-list'),
      client.requestProperty('sid'),
      client.requestProperty('sub-start'),
    ]);

    const currentStart =
      typeof subStartRaw === 'number' && Number.isFinite(subStartRaw) ? subStartRaw : null;
    if (currentStart === null) {
      throw new Error('Current subtitle start time is unavailable.');
    }

    const source = getActiveSubtitleSource(trackListRaw, sidRaw);
    let cueStarts = cueCache.get(source)?.starts;
    if (!cueStarts) {
      const content = await deps.loadSubtitleSourceText(source);
      cueStarts = parseCueStarts(content, source);
      cueCache.set(source, { starts: cueStarts });
    }

    const targetStart = findAdjacentCueStart(cueStarts, currentStart, direction);
    const delta = targetStart - currentStart;
    deps.sendMpvCommand(['add', 'sub-delay', delta]);
    deps.showMpvOsd('Subtitle delay: ${sub-delay}');
  };
}
