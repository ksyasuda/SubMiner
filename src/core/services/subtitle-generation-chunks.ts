import type { SubtitleCue } from './subtitle-cue-parser';
import { SPEECH_PASSAGE_SECONDS, type SpeechPassage } from './subtitle-generation-speech';

const CHUNK_CONTEXT_SECONDS = 0.25;
const WHISPER_WINDOW_SECONDS = 30;
const PAUSE_SEARCH_SECONDS = 5;

// Prefer detected speech starts, then quiet pauses. Context stays inside retained audio.
export function splitSpeechPassages(
  passages: readonly SpeechPassage[],
  pauses: readonly number[] = [],
  speechStarts: readonly number[] = [],
): SpeechPassage[] {
  return passages.flatMap((passage) => {
    if (passage.endSeconds - passage.startSeconds <= WHISPER_WINDOW_SECONDS)
      return [{ ...passage }];
    const chunks: SpeechPassage[] = [];
    let boundary = passage.startSeconds;
    while (boundary < passage.endSeconds) {
      const target = boundary + SPEECH_PASSAGE_SECONDS;
      let end = Math.min(target, passage.endSeconds);
      if (target < passage.endSeconds) {
        // Starting in a long quiet lead-in can make Whisper place the next line
        // several seconds early. A nearby VAD start gives the next chunk an anchor.
        let nearestSpeechStart: number | undefined;
        for (const time of speechStarts) {
          if (
            time >= target - PAUSE_SEARCH_SECONDS &&
            time <= target + PAUSE_SEARCH_SECONDS &&
            time < passage.endSeconds &&
            (nearestSpeechStart === undefined ||
              Math.abs(time - target) < Math.abs(nearestSpeechStart - target))
          )
            nearestSpeechStart = time;
        }
        let latestPause: number | undefined;
        for (const time of pauses) {
          if (
            time >= target - PAUSE_SEARCH_SECONDS &&
            time <= target &&
            (latestPause === undefined || time > latestPause)
          )
            latestPause = time;
        }
        end = nearestSpeechStart ?? latestPause ?? end;
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
    const text = cue.text.replace(/[\s\p{P}]+/gu, '');
    let duplicate: SubtitleCue | undefined;
    let greatestOverlap = 0;
    for (const [index, previous] of cues.entries()) {
      if (index >= previousCount) break;
      if (!text || matched.has(previous) || previous.text.replace(/[\s\p{P}]+/gu, '') !== text)
        continue;
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
