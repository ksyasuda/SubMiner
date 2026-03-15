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

// Stub exports — will be implemented in subsequent tasks
export function parseAssCues(_content: string): SubtitleCue[] {
  throw new Error('parseAssCues not yet implemented');
}

export function parseSubtitleCues(_content: string, _filename: string): SubtitleCue[] {
  throw new Error('parseSubtitleCues not yet implemented');
}
