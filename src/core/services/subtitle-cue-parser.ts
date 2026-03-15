export interface SubtitleCue {
  startTime: number;
  endTime: number;
  text: string;
}

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

    const startTime = parseTimestamp(timingMatch[1], timingMatch[2]!, timingMatch[3]!, timingMatch[4]!);
    const endTime = parseTimestamp(timingMatch[5], timingMatch[6]!, timingMatch[7]!, timingMatch[8]!);

    i += 1;
    const textLines: string[] = [];
    while (i < lines.length && lines[i]!.trim() !== '') {
      textLines.push(lines[i]!);
      i += 1;
    }

    const text = textLines.join('\n').trim();
    if (text) {
      cues.push({ startTime, endTime, text });
    }
  }

  return cues;
}

const ASS_OVERRIDE_TAG_PATTERN = /\{[^}]*\}/g;

const ASS_TIMING_PATTERN = /^(\d+):(\d{2}):(\d{2})\.(\d{1,2})$/;

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

  for (const line of lines) {
    const trimmed = line.trim();

    if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
      inEventsSection = trimmed.toLowerCase() === '[events]';
      continue;
    }

    if (!inEventsSection) {
      continue;
    }

    if (!trimmed.startsWith('Dialogue:')) {
      continue;
    }

    // Split on first 9 commas (ASS v4+ has 10 fields; last is Text which can contain commas)
    const afterPrefix = trimmed.slice('Dialogue:'.length);
    const fields: string[] = [];
    let remaining = afterPrefix;
    for (let fieldIndex = 0; fieldIndex < 9; fieldIndex += 1) {
      const commaIndex = remaining.indexOf(',');
      if (commaIndex < 0) {
        break;
      }
      fields.push(remaining.slice(0, commaIndex));
      remaining = remaining.slice(commaIndex + 1);
    }

    if (fields.length < 9) {
      continue;
    }

    const startTime = parseAssTimestamp(fields[1]!);
    const endTime = parseAssTimestamp(fields[2]!);
    if (startTime === null || endTime === null) {
      continue;
    }

    const rawText = remaining.replace(ASS_OVERRIDE_TAG_PATTERN, '').trim();
    if (rawText) {
      cues.push({ startTime, endTime, text: rawText });
    }
  }

  return cues;
}

// Stub export — will be implemented in Task 6
export function parseSubtitleCues(_content: string, _filename: string): SubtitleCue[] {
  throw new Error('parseSubtitleCues not yet implemented');
}
