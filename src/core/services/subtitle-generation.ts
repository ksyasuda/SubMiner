import { access, mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { constants } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import type {
  SubtitleGenerationConfig,
  SubtitleGenerationProgress,
} from '../../shared/subtitle-generation';
import { isMissingFile, resolveSubtitleGenerationModel } from './subtitle-generation-models';
import { runSubtitleGenerationProcess } from './subtitle-generation-process';
import { publishSubtitleGenerationFile } from './subtitle-generation-files';
import { formatTimestamp } from './subtitle-generation-srt';
import { transcribeSubtitleDialogue } from './subtitle-generation-dialogue';
import {
  requireSubtitleGenerationTools,
  resolveSubtitleGenerationTools,
} from './subtitle-generation-tools';

export {
  downloadSubtitleGenerationModel,
  resolveSubtitleGenerationModel,
} from './subtitle-generation-models';
export { resolveSubtitleGenerationTools } from './subtitle-generation-tools';

function numericTime(value: unknown): number | undefined {
  if (typeof value !== 'number' && typeof value !== 'string') return undefined;
  const number = Number(value);
  return Number.isFinite(number) ? number : undefined;
}

function parseAudioProbe(raw: string, selectedIndex: number | undefined) {
  const value: unknown = JSON.parse(raw);
  if (
    typeof value !== 'object' ||
    value === null ||
    !('streams' in value) ||
    !Array.isArray(value.streams)
  ) {
    throw new Error('ffprobe did not return media streams.');
  }
  const streams = value.streams.flatMap((stream: unknown) => {
    if (
      typeof stream !== 'object' ||
      stream === null ||
      !('codec_type' in stream) ||
      stream.codec_type !== 'audio' ||
      !('index' in stream) ||
      typeof stream.index !== 'number' ||
      !Number.isInteger(stream.index) ||
      stream.index < 0
    )
      return [];
    const tags = 'tags' in stream ? stream.tags : undefined;
    const language =
      typeof tags === 'object' && tags !== null && 'language' in tags ? tags.language : undefined;
    return [
      {
        index: stream.index,
        start: 'start_time' in stream ? numericTime(stream.start_time) : undefined,
        duration: 'duration' in stream ? numericTime(stream.duration) : undefined,
        japanese: language === 'ja' || language === 'jpn',
      },
    ];
  });
  const selected =
    selectedIndex === undefined
      ? (streams.find((stream) => stream.japanese) ?? streams[0])
      : streams.find((stream) => stream.index === selectedIndex);
  if (!selected)
    throw new Error(
      selectedIndex === undefined
        ? 'No audio track found.'
        : `Audio stream ${selectedIndex} was not found.`,
    );
  const format = 'format' in value ? value.format : undefined;
  const formatStart =
    typeof format === 'object' && format !== null && 'start_time' in format
      ? (numericTime(format.start_time) ?? 0)
      : 0;
  const duration =
    selected.duration ??
    (typeof format === 'object' && format !== null && 'duration' in format
      ? numericTime(format.duration)
      : undefined);
  // mpv rebases media timestamps to the container start. Extraction rebases the selected audio.
  return { index: selected.index, offset: (selected.start ?? formatStart) - formatStart, duration };
}

function shiftSubtitleTimestamps(srt: string, offsetSeconds: number): string {
  let cueCount = 0;
  const result = srt.replace(
    /(\d{2,}):(\d{2}):(\d{2}),(\d{3}) --> (\d{2,}):(\d{2}):(\d{2}),(\d{3})/g,
    (
      _match,
      sh: string,
      sm: string,
      ss: string,
      sms: string,
      eh: string,
      em: string,
      es: string,
      ems: string,
    ) => {
      cueCount += 1;
      const start = Number(sh) * 3600000 + Number(sm) * 60000 + Number(ss) * 1000 + Number(sms);
      const end = Number(eh) * 3600000 + Number(em) * 60000 + Number(es) * 1000 + Number(ems);
      return `${formatTimestamp(start + offsetSeconds * 1000)} --> ${formatTimestamp(end + offsetSeconds * 1000)}`;
    },
  );
  if (cueCount === 0)
    throw new Error(
      'Whisper produced no subtitle cues. The audio may contain no recognized speech.',
    );
  return result;
}

async function ensureAvailableOutput(outputPath: string): Promise<void> {
  try {
    await stat(outputPath);
  } catch (error) {
    if (isMissingFile(error)) return;
    throw error;
  }
  throw new Error(`Subtitle output already exists: ${outputPath}`);
}

// Fail before extraction and transcription when the destination cannot take the file.
export async function ensureWritableDirectory(directory: string): Promise<void> {
  try {
    await access(directory, constants.W_OK | constants.X_OK);
  } catch {
    throw new Error(`Cannot save subtitles: ${directory} is not writable.`);
  }
}

async function writeSubtitles(input: {
  mediaPath: string;
  outputPath?: string;
  contents: string;
  signal?: AbortSignal;
}): Promise<string> {
  const parsed = path.parse(input.mediaPath);
  const directory = input.outputPath ? path.dirname(path.resolve(input.outputPath)) : parsed.dir;
  const temporaryDirectory = await mkdtemp(path.join(directory, '.subminer-subtitles-'));
  try {
    const staged = path.join(temporaryDirectory, 'subtitles.srt');
    await writeFile(staged, input.contents, { flag: 'wx' });
    for (let suffix = 0; ; suffix += 1) {
      input.signal?.throwIfAborted();
      const destination = input.outputPath
        ? path.resolve(input.outputPath)
        : path.join(directory, `${parsed.name}.ja.generated${suffix ? `.${suffix}` : ''}.srt`);
      try {
        await publishSubtitleGenerationFile(staged, destination);
        return destination;
      } catch (error) {
        if (
          !input.outputPath &&
          error instanceof Error &&
          'code' in error &&
          error.code === 'EEXIST'
        )
          continue;
        throw error;
      }
    }
  } finally {
    await rm(temporaryDirectory, { recursive: true, force: true });
  }
}

export async function generateJapaneseSubtitles(input: {
  config: SubtitleGenerationConfig;
  modelDirectory: string;
  mediaPath: string;
  audioStreamIndex?: number;
  outputPath?: string;
  onProgress?: (progress: SubtitleGenerationProgress) => void;
  signal?: AbortSignal;
}): Promise<string> {
  input.signal?.throwIfAborted();
  if (/^[a-z][a-z\d+.-]*:\/\//i.test(input.mediaPath))
    throw new Error('Subtitle generation requires a local media file.');
  const mediaPath = path.resolve(input.mediaPath);
  if (!(await stat(mediaPath)).isFile())
    throw new Error('Subtitle generation requires a local media file.');
  if (input.outputPath) await ensureAvailableOutput(path.resolve(input.outputPath));
  await ensureWritableDirectory(
    input.outputPath ? path.dirname(path.resolve(input.outputPath)) : path.dirname(mediaPath),
  );
  const model = await resolveSubtitleGenerationModel(input.config, input.modelDirectory);
  if (model.kind === 'missing')
    throw new Error(
      'No Whisper model found. Download a model or configure an existing model path.',
    );
  if (model.kind === 'invalid') throw new Error(model.message);
  const tools = requireSubtitleGenerationTools(await resolveSubtitleGenerationTools(input.config));
  input.onProgress?.({ stage: 'extract', message: 'Inspecting audio tracks...' });
  const probe = await runSubtitleGenerationProcess({
    command: tools.ffprobe,
    args: [
      '-v',
      'error',
      '-show_entries',
      'stream=index,codec_type,start_time,duration:stream_tags=language:format=start_time,duration',
      '-of',
      'json',
      mediaPath,
    ],
    signal: input.signal,
  });
  const audio = parseAudioProbe(probe, input.audioStreamIndex);
  const temporaryDirectory = await mkdtemp(path.join(tmpdir(), 'subminer-whisper-'));
  try {
    const wavPath = path.join(temporaryDirectory, 'audio.wav');
    const subtitleBase = path.join(temporaryDirectory, 'subtitles');
    input.onProgress?.({ stage: 'extract', percent: 0, message: 'Extracting audio...' });
    await runSubtitleGenerationProcess({
      command: tools.ffmpeg,
      args: [
        '-nostdin',
        '-hide_banner',
        '-loglevel',
        'error',
        '-i',
        mediaPath,
        '-map',
        `0:${audio.index}`,
        '-vn',
        '-af',
        'asetpts=PTS-STARTPTS',
        '-ac',
        '1',
        '-ar',
        '16000',
        '-c:a',
        'pcm_s16le',
        '-progress',
        'pipe:1',
        '-nostats',
        wavPath,
      ],
      signal: input.signal,
      onLine: (line) => {
        const match = /^out_time_us=(\d+)$/.exec(line);
        if (match && audio.duration && audio.duration > 0) {
          input.onProgress?.({
            stage: 'extract',
            percent: Math.min(100, Math.floor(Number(match[1]) / 10000 / audio.duration)),
            message: 'Extracting audio...',
          });
        }
      },
    });
    input.onProgress?.({
      stage: 'transcribe',
      percent: 0,
      message: 'Generating Japanese subtitles...',
    });
    let srt: string;
    if (tools.vad !== null) {
      srt = await transcribeSubtitleDialogue({
        config: input.config,
        tools: { ...tools, vad: tools.vad },
        modelPath: model.path,
        wavPath,
        directory: temporaryDirectory,
        signal: input.signal,
        onProgress: input.onProgress,
      });
    } else {
      await runSubtitleGenerationProcess({
        command: tools.whisper,
        args: [
          '-m',
          model.path,
          '-f',
          wavPath,
          '-l',
          'ja',
          '-t',
          String(input.config.threads),
          '-osrt',
          '-of',
          subtitleBase,
          '-pp',
        ],
        signal: input.signal,
        onLine: (line) => {
          const match = /progress\s*=\s*(\d+(?:\.\d+)?)%/.exec(line);
          if (match)
            input.onProgress?.({
              stage: 'transcribe',
              percent: Math.min(100, Number(match[1])),
              message: 'Generating Japanese subtitles...',
            });
        },
      });
      srt = await readFile(`${subtitleBase}.srt`, 'utf8');
    }
    input.signal?.throwIfAborted();
    input.onProgress?.({ stage: 'write', message: 'Saving Japanese subtitles...' });
    const contents = shiftSubtitleTimestamps(srt, audio.offset);
    const outputPath = await writeSubtitles({
      mediaPath,
      outputPath: input.outputPath,
      contents,
      signal: input.signal,
    });
    input.onProgress?.({ stage: 'write', percent: 100, message: 'Japanese subtitles are ready.' });
    return outputPath;
  } finally {
    await rm(temporaryDirectory, { recursive: true, force: true });
  }
}
