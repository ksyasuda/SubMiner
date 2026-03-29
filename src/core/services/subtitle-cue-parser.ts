export interface SubtitleCue {
  startTime: number;
  endTime: number;
  text: string;
}

const HTML_SUBTITLE_TAG_PATTERN = /<\/?[A-Za-z][^>\n]*>/g;

const SRT_TIMING_PATTERN =
  /^\s*(?:(\d{1,2}):)?(\d{2}):(\d{2})[,.](\d{1,3})\s*-->\s*(?:(\d{1,2}):)?(\d{2}):(\d{2})[,.](\d{1,3})/;

function parseTimestamp(
  hours: string | undefined,
  minutes: string,
  seconds: string,
  millis: string,
): number {
  return (
    Number(hours || 0) * 3600 +
    Number(minutes) * 60 +
    Number(seconds) +
    Number(millis.padEnd(3, '0')) / 1000
  );
}

function sanitizeSubtitleCueText(text: string): string {
  return text.replace(ASS_OVERRIDE_TAG_PATTERN, '').replace(HTML_SUBTITLE_TAG_PATTERN, '').trim();
}

export function parseSrtCues(content: string): SubtitleCue[] {
  const cues: SubtitleCue[] = [];
  const lines = content.split(/\r?\n/);
  let i = 0;

  while (i < lines.length) {
    const line = lines[i]!;
    const timingMatch = SRT_TIMING_PATTERN.exec(line);
    if (!timingMatch) {
      i += 1;
      continue;
    }

    const startTime = parseTimestamp(
      timingMatch[1],
      timingMatch[2]!,
      timingMatch[3]!,
      timingMatch[4]!,
    );
    const endTime = parseTimestamp(
      timingMatch[5],
      timingMatch[6]!,
      timingMatch[7]!,
      timingMatch[8]!,
    );

    i += 1;
    const textLines: string[] = [];
    while (i < lines.length && lines[i]!.trim() !== '') {
      textLines.push(lines[i]!);
      i += 1;
    }

    const text = sanitizeSubtitleCueText(textLines.join('\n'));
    if (text) {
      cues.push({ startTime, endTime, text });
    }
  }

  return cues;
}

const ASS_OVERRIDE_TAG_PATTERN = /\{[^}]*\}/g;

const ASS_TIMING_PATTERN = /^(\d+):(\d{2}):(\d{2})\.(\d{1,2})$/;
const ASS_FORMAT_PREFIX = 'Format:';
const ASS_DIALOGUE_PREFIX = 'Dialogue:';

function parseAssTimestamp(raw: string): number | null {
  const match = ASS_TIMING_PATTERN.exec(raw.trim());
  if (!match) {
    return null;
  }
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  const seconds = Number(match[3]);
  const centiseconds = Number(match[4]!.padEnd(2, '0'));
  return hours * 3600 + minutes * 60 + seconds + centiseconds / 100;
}

export function parseAssCues(content: string): SubtitleCue[] {
  const cues: SubtitleCue[] = [];
  const lines = content.split(/\r?\n/);
  let inEventsSection = false;
  let startFieldIndex = -1;
  let endFieldIndex = -1;
  let textFieldIndex = -1;

  for (const line of lines) {
    const trimmed = line.trim();

    if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
      inEventsSection = trimmed.toLowerCase() === '[events]';
      if (!inEventsSection) {
        startFieldIndex = -1;
        endFieldIndex = -1;
        textFieldIndex = -1;
      }
      continue;
    }

    if (!inEventsSection) {
      continue;
    }

    if (trimmed.startsWith(ASS_FORMAT_PREFIX)) {
      const formatFields = trimmed
        .slice(ASS_FORMAT_PREFIX.length)
        .split(',')
        .map((field) => field.trim().toLowerCase());
      startFieldIndex = formatFields.indexOf('start');
      endFieldIndex = formatFields.indexOf('end');
      textFieldIndex = formatFields.indexOf('text');
      continue;
    }

    if (!trimmed.startsWith(ASS_DIALOGUE_PREFIX)) {
      continue;
    }

    if (startFieldIndex < 0 || endFieldIndex < 0 || textFieldIndex < 0) {
      continue;
    }

    const fields = trimmed.slice(ASS_DIALOGUE_PREFIX.length).split(',');
    if (
      startFieldIndex >= fields.length ||
      endFieldIndex >= fields.length ||
      textFieldIndex >= fields.length
    ) {
      continue;
    }

    const startTime = parseAssTimestamp(fields[startFieldIndex]!);
    const endTime = parseAssTimestamp(fields[endFieldIndex]!);
    if (startTime === null || endTime === null) {
      continue;
    }

    const text = sanitizeSubtitleCueText(fields.slice(textFieldIndex).join(','));
    if (text) {
      cues.push({ startTime, endTime, text });
    }
  }

  return cues;
}

function detectSubtitleFormat(source: string): 'srt' | 'vtt' | 'ass' | 'ssa' | null {
  const [normalizedSource = source] =
    (() => {
      try {
        return /^[a-z]+:\/\//i.test(source) ? new URL(source).pathname : source;
      } catch {
        return source;
      }
    })().split(/[?#]/, 1)[0] ?? '';
  const ext = normalizedSource.split('.').pop()?.toLowerCase() ?? '';
  if (ext === 'srt') return 'srt';
  if (ext === 'vtt') return 'vtt';
  if (ext === 'ass' || ext === 'ssa') return 'ass';
  return null;
}

export function parseSubtitleCues(content: string, filename: string): SubtitleCue[] {
  const format = detectSubtitleFormat(filename);
  let cues: SubtitleCue[];

  switch (format) {
    case 'srt':
    case 'vtt':
      cues = parseSrtCues(content);
      break;
    case 'ass':
    case 'ssa':
      cues = parseAssCues(content);
      break;
    default:
      cues = [];
  }

  if (cues.length === 0) {
    const assCues = parseAssCues(content);
    const srtCues = parseSrtCues(content);
    cues = assCues.length >= srtCues.length ? assCues : srtCues;
  }

  cues.sort((a, b) => a.startTime - b.startTime);
  return cues;
}
