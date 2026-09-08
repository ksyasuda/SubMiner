import { parseSrtCues, type SubtitleCue } from './subtitle-cue-parser';

export interface SpeechPassage {
  startSeconds: number;
  endSeconds: number;
}

export const SPEECH_PASSAGE_SECONDS = 20;

// The standalone whisper.cpp detector reports centiseconds, unlike its diagnostic logs.
export function parseSpeechPassages(output: string): SpeechPassage[] {
  const count = /^Detected (\d+) speech segments:$/m.exec(output);
  if (!count) throw new Error('Speech detector did not report its segment count.');
  const passages: SpeechPassage[] = [];
  for (const line of output.split(/\r?\n/)) {
    if (!line.startsWith('Speech segment ')) continue;
    const match = /^Speech segment (\d+): start = (\d+(?:\.\d+)?), end = (\d+(?:\.\d+)?)$/.exec(
      line,
    );
    if (!match) throw new Error('Speech detector returned a malformed segment.');
    const index = Number(match[1]);
    const startSeconds = Number(match[2]) / 100;
    const endSeconds = Number(match[3]) / 100;
    if (
      index !== passages.length ||
      !Number.isFinite(startSeconds) ||
      !Number.isFinite(endSeconds) ||
      endSeconds <= startSeconds ||
      startSeconds < (passages.at(-1)?.endSeconds ?? 0) ||
      // VAD can pass the requested split point while looking for a pause.
      endSeconds - startSeconds > 30
    ) {
      throw new Error('Speech detector returned unordered or invalid segment timing.');
    }
    passages.push({ startSeconds, endSeconds });
  }
  if (passages.length !== Number(count[1]))
    throw new Error('Speech detector output is incomplete.');

  const grouped: SpeechPassage[] = [];
  for (const passage of passages) {
    const previous = grouped.at(-1);
    if (
      previous &&
      passage.startSeconds - previous.endSeconds <= 1 &&
      passage.endSeconds - previous.startSeconds <= SPEECH_PASSAGE_SECONDS
    ) {
      previous.endSeconds = passage.endSeconds;
    } else grouped.push({ ...passage });
  }
  return grouped;
}

// Clamp to the audio actually supplied to Whisper. A cue cannot cross an omitted music break.
export function speechPassageCues(srt: string, passage: SpeechPassage): SubtitleCue[] {
  const duration = passage.endSeconds - passage.startSeconds;
  const cues = parseSrtCues(srt);
  if (srt.trim() && cues.length === 0)
    throw new Error('Whisper returned malformed subtitles for a speech passage.');
  return cues.flatMap((cue) => {
    const start = Math.max(0, cue.startTime);
    const end = Math.min(duration, cue.endTime);
    if (!Number.isFinite(start) || !Number.isFinite(end) || end <= start) return [];
    return [
      { ...cue, startTime: passage.startSeconds + start, endTime: passage.startSeconds + end },
    ];
  });
}
