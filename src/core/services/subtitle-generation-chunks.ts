import type { SubtitleCue } from './subtitle-cue-parser';
import { SPEECH_PASSAGE_SECONDS, type SpeechPassage } from './subtitle-generation-speech';

const CHUNK_CONTEXT_SECONDS = 0.25;
const PAUSE_SEARCH_SECONDS = 5;

// Prefer a quiet pause near the end of each chunk. Context stays inside detected speech.
export function splitSpeechPassages(
  passages: readonly SpeechPassage[],
  pauses: readonly number[] = [],
): SpeechPassage[] {
  return passages.flatMap((passage) => {
    const chunks: SpeechPassage[] = [];
    let boundary = passage.startSeconds;
    while (boundary < passage.endSeconds) {
      const target = boundary + SPEECH_PASSAGE_SECONDS;
      let end = Math.min(target, passage.endSeconds);
      if (target < passage.endSeconds) {
        let latestPause: number | undefined;
        for (const time of pauses) {
          if (
            time >= target - PAUSE_SEARCH_SECONDS &&
            time <= target &&
            (latestPause === undefined || time > latestPause)
          )
            latestPause = time;
        }
        if (latestPause !== undefined) end = latestPause;
      }
      chunks.push({
        startSeconds: Math.max(passage.startSeconds, boundary - CHUNK_CONTEXT_SECONDS),
        endSeconds: Math.min(passage.endSeconds, end + CHUNK_CONTEXT_SECONDS),
      });
      boundary = end;
    }
    return chunks;
  });
}

// Deduplicate only matching text substantially overlapping cues from earlier chunks.
// Repeated words within the current chunk or at separate times remain separate.
export function appendSpeechChunkCues(cues: SubtitleCue[], incoming: readonly SubtitleCue[]): void {
  const previousCount = cues.length;
  const matched = new Set<SubtitleCue>();
  for (const cue of incoming) {
    const text = cue.text.replace(/\s+/g, '');
    let duplicate: SubtitleCue | undefined;
    let greatestOverlap = 0;
    for (const [index, previous] of cues.entries()) {
      if (index >= previousCount) break;
      if (matched.has(previous) || previous.text.replace(/\s+/g, '') !== text) continue;
      const overlap =
        Math.min(previous.endTime, cue.endTime) - Math.max(previous.startTime, cue.startTime);
      const shorterDuration = Math.min(
        previous.endTime - previous.startTime,
        cue.endTime - cue.startTime,
      );
      if (overlap > greatestOverlap && overlap >= shorterDuration / 2) {
        duplicate = previous;
        greatestOverlap = overlap;
      }
    }
    if (duplicate) {
      duplicate.startTime = Math.min(duplicate.startTime, cue.startTime);
      duplicate.endTime = Math.max(duplicate.endTime, cue.endTime);
      matched.add(duplicate);
    } else cues.push({ ...cue });
  }
}
