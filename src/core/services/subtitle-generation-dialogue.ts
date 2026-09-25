import { access, readFile, rm } from 'node:fs/promises';
import { constants } from 'node:fs';
import path from 'node:path';
import type {
  SubtitleGenerationConfig,
  SubtitleGenerationProgress,
} from '../../shared/subtitle-generation';
import { expandSubtitleGenerationPath } from './subtitle-generation-files';
import { runSubtitleGenerationProcess } from './subtitle-generation-process';
import type { SubtitleGenerationToolPaths } from './subtitle-generation-tools';
import { formatTimestamp } from './subtitle-generation-srt';
import {
  parseSpeechPassages,
  speechPassageCues,
  SPEECH_PASSAGE_SECONDS,
  type SpeechPassage,
} from './subtitle-generation-speech';
import type { SubtitleCue } from './subtitle-cue-parser';
import { appendSpeechChunkCues, splitSpeechPassages } from './subtitle-generation-chunks';
import { findSpeechPauses } from './subtitle-generation-pauses';
import { findAudiblePassages, mergeSpeechPassages } from './subtitle-generation-coverage';

export async function transcribeSubtitleDialogue(input: {
  config: SubtitleGenerationConfig;
  tools: SubtitleGenerationToolPaths;
  referenceStarts?: readonly number[];
  modelPath: string;
  wavPath: string;
  directory: string;
  onProgress?: (progress: SubtitleGenerationProgress) => void;
  signal?: AbortSignal;
}): Promise<string> {
  let speech: SpeechPassage[] = [];
  if (input.tools.vad !== null) {
    const vadModelPath = expandSubtitleGenerationPath(input.config.vadModelPath);
    await access(vadModelPath, constants.R_OK);
    input.onProgress?.({ stage: 'transcribe', percent: 0, message: 'Finding spoken dialogue...' });
    const segmentLines: string[] = [];
    await runSubtitleGenerationProcess({
      command: input.tools.vad,
      args: [
        '-f',
        input.wavPath,
        '-vm',
        vadModelPath,
        '-t',
        String(input.config.threads),
        '-vt',
        '0.3',
        '--vad-min-speech-duration-ms',
        '100',
        '--vad-min-silence-duration-ms',
        '500',
        '-vp',
        '350',
        '-vmsd',
        String(SPEECH_PASSAGE_SECONDS),
        '-np',
      ],
      signal: input.signal,
      // Capture structured result lines separately from the bounded process log.
      onLine: (line) => {
        if (line.startsWith('Detected ') || line.startsWith('Speech segment '))
          segmentLines.push(line);
      },
    });
    speech = parseSpeechPassages(segmentLines.join('\n'));
  }
  input.onProgress?.({ stage: 'transcribe', percent: 0, message: 'Checking audio coverage...' });
  const audible = await findAudiblePassages({
    ffmpegPath: input.tools.ffmpeg,
    wavPath: input.wavPath,
    signal: input.signal,
  });
  const detected = mergeSpeechPassages([...speech, ...audible]);
  if (detected.length === 0) throw new Error('No spoken dialogue detected.');
  const pauses = detected.some(
    (passage) => passage.endSeconds - passage.startSeconds > SPEECH_PASSAGE_SECONDS,
  )
    ? await findSpeechPauses({
        ffmpegPath: input.tools.ffmpeg,
        wavPath: input.wavPath,
        signal: input.signal,
      })
    : [];
  const passages = splitSpeechPassages(
    detected,
    pauses,
    speech.map((passage) => passage.startSeconds),
    input.referenceStarts,
  );
  const cues: SubtitleCue[] = [];
  for (const [index, passage] of passages.entries()) {
    const base = path.join(input.directory, `speech-${index}`);
    input.onProgress?.({
      stage: 'transcribe',
      percent: Math.floor((index / passages.length) * 100),
      message: `Transcribing dialogue passage ${index + 1} of ${passages.length}...`,
    });
    await runSubtitleGenerationProcess({
      command: input.tools.ffmpeg,
      args: [
        '-nostdin',
        '-hide_banner',
        '-loglevel',
        'error',
        '-ss',
        String(passage.startSeconds),
        '-i',
        input.wavPath,
        '-t',
        String(passage.endSeconds - passage.startSeconds),
        '-ac',
        '1',
        '-ar',
        '16000',
        '-c:a',
        'pcm_s16le',
        `${base}.wav`,
      ],
      signal: input.signal,
    });
    // -mc 0 limits text context, but does not isolate decoder state across input files.
    // A fresh process prevents earlier passages from corrupting later transcriptions.
    await runSubtitleGenerationProcess({
      command: input.tools.whisper,
      args: [
        '-m',
        input.modelPath,
        '-l',
        'ja',
        '-t',
        String(input.config.threads),
        '-mc',
        '0',
        '-sns',
        '-osrt',
        '-f',
        `${base}.wav`,
        '-of',
        base,
      ],
      signal: input.signal,
    });
    input.signal?.throwIfAborted();
    appendSpeechChunkCues(cues, speechPassageCues(await readFile(`${base}.srt`, 'utf8'), passage));
    await rm(`${base}.wav`);
  }
  if (cues.length === 0) throw new Error('Whisper recognized no dialogue in the detected speech.');
  return cues
    .sort((a, b) => a.startTime - b.startTime || a.endTime - b.endTime)
    .map(
      (cue, index) =>
        `${index + 1}\n${formatTimestamp(cue.startTime * 1000)} --> ${formatTimestamp(cue.endTime * 1000)}\n${cue.text}\n`,
    )
    .join('\n');
}
