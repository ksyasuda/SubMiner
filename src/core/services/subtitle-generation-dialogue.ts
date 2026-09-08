import { access, readFile, rm } from 'node:fs/promises';
import { constants } from 'node:fs';
import path from 'node:path';
import type {
  SubtitleGenerationConfig,
  SubtitleGenerationProgress,
} from '../../shared/subtitle-generation';
import { expandSubtitleGenerationPath } from './subtitle-generation-files';
import { runSubtitleGenerationProcess } from './subtitle-generation-process';
import { formatTimestamp } from './subtitle-generation-srt';
import {
  parseSpeechPassages,
  speechPassageCues,
  SPEECH_PASSAGE_SECONDS,
} from './subtitle-generation-speech';
import type { SubtitleCue } from './subtitle-cue-parser';

const PASSAGES_PER_BATCH = 16;

export async function transcribeSubtitleDialogue(input: {
  config: SubtitleGenerationConfig;
  modelPath: string;
  wavPath: string;
  directory: string;
  onProgress?: (progress: SubtitleGenerationProgress) => void;
  signal?: AbortSignal;
}): Promise<string> {
  const vadModelPath = expandSubtitleGenerationPath(input.config.vadModelPath);
  await access(vadModelPath, constants.R_OK);
  input.onProgress?.({ stage: 'transcribe', percent: 0, message: 'Finding spoken dialogue...' });
  const segmentLines: string[] = [];
  await runSubtitleGenerationProcess({
    command: input.config.vadPath.trim() || 'whisper-vad-speech-segments',
    args: [
      '-f',
      input.wavPath,
      '-vm',
      vadModelPath,
      '-t',
      String(input.config.threads),
      '-vt',
      '0.3',
      '-vsd',
      '500',
      '-vp',
      '200',
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
  const passages = parseSpeechPassages(segmentLines.join('\n'));
  if (passages.length === 0) throw new Error('No spoken dialogue detected.');
  const cues: SubtitleCue[] = [];
  for (let offset = 0; offset < passages.length; offset += PASSAGES_PER_BATCH) {
    const batch = passages.slice(offset, offset + PASSAGES_PER_BATCH).map((passage, index) => ({
      passage,
      base: path.join(input.directory, `speech-${offset + index}`),
    }));
    input.onProgress?.({
      stage: 'transcribe',
      percent: Math.floor((offset / passages.length) * 100),
      message: `Transcribing dialogue passages ${offset + 1}-${offset + batch.length} of ${passages.length}...`,
    });
    for (const { passage, base } of batch) {
      await runSubtitleGenerationProcess({
        command: input.config.ffmpegPath.trim() || 'ffmpeg',
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
    }
    // One model load per batch, with independent text context and timestamps for every passage.
    await runSubtitleGenerationProcess({
      command: input.config.whisperPath.trim() || 'whisper-cli',
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
        ...batch.flatMap(({ base }) => ['-f', `${base}.wav`, '-of', base]),
      ],
      signal: input.signal,
    });
    for (const { passage, base } of batch) {
      input.signal?.throwIfAborted();
      cues.push(...speechPassageCues(await readFile(`${base}.srt`, 'utf8'), passage));
      await rm(`${base}.wav`);
    }
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
