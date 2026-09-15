import { runSubtitleGenerationProcess } from './subtitle-generation-process';
import type { SpeechPassage } from './subtitle-generation-speech';

const AUDIO_PADDING_SECONDS = 0.35;

export function mergeSpeechPassages(passages: readonly SpeechPassage[]): SpeechPassage[] {
  const merged: SpeechPassage[] = [];
  for (const passage of [...passages].sort((a, b) => a.startSeconds - b.startSeconds)) {
    const previous = merged.at(-1);
    if (previous && passage.startSeconds <= previous.endSeconds)
      previous.endSeconds = Math.max(previous.endSeconds, passage.endSeconds);
    else merged.push({ ...passage });
  }
  return merged;
}

// VAD rejection is not proof of silence. Preserve audible gaps for Whisper to evaluate.
export async function findAudiblePassages(input: {
  ffmpegPath: string;
  wavPath: string;
  signal?: AbortSignal;
}): Promise<SpeechPassage[]> {
  const silences: SpeechPassage[] = [];
  let duration = 0;
  let trailingSilence: number | undefined;
  await runSubtitleGenerationProcess({
    command: input.ffmpegPath,
    args: [
      '-nostdin',
      '-hide_banner',
      '-nostats',
      '-i',
      input.wavPath,
      '-af',
      'silencedetect=noise=-50dB:d=0.5',
      '-progress',
      'pipe:1',
      '-f',
      'null',
      '-',
    ],
    signal: input.signal,
    onLine: (line) => {
      const progress = /^out_time_us=(\d+)$/.exec(line);
      if (progress) duration = Math.max(duration, Number(progress[1]) / 1_000_000);
      const start = /silence_start: (\S+)/.exec(line);
      if (start && Number.isFinite(Number(start[1]))) trailingSilence = Number(start[1]);
      const end = /silence_end: (\S+) \| silence_duration: (\S+)/.exec(line);
      if (!end) return;
      const endSeconds = Number(end[1]);
      const length = Number(end[2]);
      if (Number.isFinite(endSeconds) && Number.isFinite(length) && length > 0) {
        silences.push({ startSeconds: Math.max(0, endSeconds - length), endSeconds });
        trailingSilence = undefined;
      }
    },
  });
  if (!Number.isFinite(duration) || duration <= 0)
    throw new Error('Audio analysis did not report a valid duration.');
  if (trailingSilence !== undefined)
    silences.push({ startSeconds: trailingSilence, endSeconds: duration });

  const audible: SpeechPassage[] = [];
  let cursor = 0;
  for (const silence of mergeSpeechPassages(silences)) {
    if (cursor >= duration) break;
    if (silence.startSeconds > cursor)
      audible.push({ startSeconds: cursor, endSeconds: Math.min(duration, silence.startSeconds) });
    cursor = Math.max(cursor, silence.endSeconds);
  }
  if (cursor < duration) audible.push({ startSeconds: cursor, endSeconds: duration });
  return mergeSpeechPassages(
    audible.map((passage) => ({
      startSeconds: Math.max(0, passage.startSeconds - AUDIO_PADDING_SECONDS),
      endSeconds: Math.min(duration, passage.endSeconds + AUDIO_PADDING_SECONDS),
    })),
  );
}
