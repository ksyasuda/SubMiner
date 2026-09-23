import { access, mkdir, mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { constants } from 'node:fs';
import { createHash } from 'node:crypto';
import { tmpdir } from 'node:os';
import path from 'node:path';
import type {
  SubtitleGenerationConfig,
  SubtitleGenerationProgress,
} from '../../shared/subtitle-generation';
import { isMissingFile, resolveSubtitleGenerationModel } from './subtitle-generation-models';
import { runSubtitleGenerationProcess } from './subtitle-generation-process';
import { prepareSubtitleGenerationAudio } from './subtitle-generation-audio';
import { publishSubtitleGenerationFile } from './subtitle-generation-files';
import { formatTimestamp } from './subtitle-generation-srt';
import { transcribeSubtitleDialogue } from './subtitle-generation-dialogue';
import { type SubtitleGenerationRemoteSource } from './subtitle-generation-source';
import {
  loadSubtitleGenerationReference,
  type SubtitleGenerationReference,
} from './subtitle-generation-reference';
import {
  requireSubtitleGenerationTools,
  resolveSubtitleGenerationTools,
} from './subtitle-generation-tools';

export {
  downloadSubtitleGenerationModel,
  resolveSubtitleGenerationModel,
} from './subtitle-generation-models';
export { resolveSubtitleGenerationTools } from './subtitle-generation-tools';

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
  remote?: SubtitleGenerationRemoteSource;
  references?: readonly SubtitleGenerationReference[];
  outputPath?: string;
  onProgress?: (progress: SubtitleGenerationProgress) => void;
  signal?: AbortSignal;
}): Promise<string> {
  input.signal?.throwIfAborted();
  const isRemote = /^[a-z][a-z\d+.-]*:\/\//i.test(input.mediaPath);
  if (isRemote && (!input.remote || !/^https?:\/\//i.test(input.mediaPath)))
    throw new Error('Subtitle generation requires a local media file or a supported HTTP stream.');
  if (!isRemote && input.remote) throw new Error('Expected an HTTP stream for remote generation.');
  const mediaPath = isRemote ? new URL(input.mediaPath).href : path.resolve(input.mediaPath);
  if (!isRemote && !(await stat(mediaPath)).isFile())
    throw new Error('Subtitle generation requires a local media file.');
  // URLs may contain credentials or expiring tokens. Keep them out of cache filenames.
  const destinationMediaPath = input.remote
    ? path.join(
        input.remote.cacheDirectory,
        createHash('sha256').update(mediaPath).digest('hex').slice(0, 24),
      )
    : mediaPath;
  if (input.remote) await mkdir(input.remote.cacheDirectory, { recursive: true });
  if (input.outputPath) await ensureAvailableOutput(path.resolve(input.outputPath));
  await ensureWritableDirectory(
    input.outputPath
      ? path.dirname(path.resolve(input.outputPath))
      : path.dirname(destinationMediaPath),
  );
  const model = await resolveSubtitleGenerationModel(input.config, input.modelDirectory);
  if (model.kind === 'missing')
    throw new Error(
      'No Whisper model found. Download a model or configure an existing model path.',
    );
  if (model.kind === 'invalid') throw new Error(model.message);
  const tools = requireSubtitleGenerationTools(await resolveSubtitleGenerationTools(input.config));
  const temporaryDirectory = await mkdtemp(
    path.join(input.remote?.sessionDirectory ?? tmpdir(), 'subminer-whisper-'),
  );
  try {
    const subtitleBase = path.join(temporaryDirectory, 'subtitles');
    const audio = await prepareSubtitleGenerationAudio({
      mediaPath,
      audioStreamIndex: input.audioStreamIndex,
      remote: input.remote,
      ffprobe: tools.ffprobe,
      ffmpeg: tools.ffmpeg,
      directory: temporaryDirectory,
      onProgress: input.onProgress,
      signal: input.signal,
    });
    const { wavPath } = audio;
    input.onProgress?.({
      stage: 'transcribe',
      percent: 0,
      message: 'Generating Japanese subtitles...',
    });
    const referenceStarts = await loadSubtitleGenerationReference({
      references: input.references ?? [],
      mediaPath,
      ffmpegPath: tools.ffmpeg,
      directory: temporaryDirectory,
      audioOffset: audio.offset,
      httpHeaders: input.remote?.httpHeaders,
      onProgress: input.onProgress,
      signal: input.signal,
    });
    let srt: string;
    if (tools.vad !== null || referenceStarts.length > 0) {
      srt = await transcribeSubtitleDialogue({
        config: input.config,
        tools,
        referenceStarts,
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
      mediaPath: destinationMediaPath,
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
