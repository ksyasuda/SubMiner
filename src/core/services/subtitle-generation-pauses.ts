import { runSubtitleGenerationProcess } from './subtitle-generation-process';

// Each completed silencedetect line contains both the end and duration of a quiet interval.
export async function findSpeechPauses(input: {
  ffmpegPath: string;
  wavPath: string;
  signal?: AbortSignal;
}): Promise<number[]> {
  const pauses: number[] = [];
  await runSubtitleGenerationProcess({
    command: input.ffmpegPath,
    args: [
      '-nostdin',
      '-hide_banner',
      '-nostats',
      '-i',
      input.wavPath,
      '-af',
      'silencedetect=noise=-35dB:d=0.12',
      '-f',
      'null',
      '-',
    ],
    signal: input.signal,
    onLine: (line) => {
      const match = /silence_end: (\S+) \| silence_duration: (\S+)/.exec(line);
      if (!match) return;
      const end = Number(match[1]);
      const duration = Number(match[2]);
      if (Number.isFinite(end) && Number.isFinite(duration) && duration > 0 && end >= duration)
        pauses.push(end - duration / 2);
    },
  });
  return pauses.sort((a, b) => a - b);
}
