export interface SrtCue {
  index: number;
  start: string;
  end: string;
  text: string;
}

const TIMING_LINE_PATTERN =
  /^(?<start>\d{2}:\d{2}:\d{2},\d{3}) --> (?<end>\d{2}:\d{2}:\d{2},\d{3})$/;

export function parseSrt(content: string): SrtCue[] {
  const normalized = content.replace(/\r\n/g, '\n').trim();
  if (!normalized) return [];

  return normalized
    .split(/\n{2,}/)
    .map((block) => {
      const lines = block.split('\n');
      const index = Number.parseInt(lines[0] || '', 10);
      const timingLine = lines[1] || '';
      const timingMatch = TIMING_LINE_PATTERN.exec(timingLine);
      if (!Number.isInteger(index) || !timingMatch?.groups) {
        throw new Error(`Invalid SRT cue block: ${block}`);
      }
      return {
        index,
        start: timingMatch.groups.start!,
        end: timingMatch.groups.end!,
        text: lines.slice(2).join('\n').trim(),
      } satisfies SrtCue;
    })
    .filter((cue) => cue.text.length > 0);
}

export function stringifySrt(cues: SrtCue[]): string {
  return cues
    .map((cue, idx) => `${idx + 1}\n${cue.start} --> ${cue.end}\n${cue.text.trim()}\n`)
    .join('\n')
    .trimEnd();
}
