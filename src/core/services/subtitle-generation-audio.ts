import path from 'node:path';
import { rm } from 'node:fs/promises';
import { probeSubtitleGenerationAudio } from './subtitle-generation-probe';
import { runSubtitleGenerationProcess } from './subtitle-generation-process';
import { downloadSubtitleGenerationHls } from './subtitle-generation-hls';
import { selectSubtitleGenerationAlternative } from './subtitle-generation-alternative';
import {
  readSubtitleAudioCache,
  writeSubtitleAudioCache,
  subtitleAudioCacheKey,
} from './subtitle-generation-audio-cache';
import {
  subtitleGenerationHttpArgs,
  type SubtitleGenerationRemoteSource,
} from './subtitle-generation-source';
import type { SubtitleGenerationProgress } from '../../shared/subtitle-generation';

export async function prepareSubtitleGenerationAudio(input: {
  mediaPath: string;
  audioStreamIndex?: number;
  remote?: SubtitleGenerationRemoteSource;
  ffprobe: string;
  ffmpeg: string;
  directory: string;
  onProgress?: (progress: SubtitleGenerationProgress) => void;
  signal?: AbortSignal;
}): Promise<{ wavPath: string; offset: number }> {
  const wavPath = path.join(input.directory, 'audio.wav');
  const selectedAudio = input.remote?.selectedAudio;
  const mediaPath = selectedAudio?.url ?? input.mediaPath;
  const audioStreamIndex = selectedAudio ? selectedAudio.audioStreamIndex : input.audioStreamIndex;
  const httpHeaders = selectedAudio?.httpHeaders ?? input.remote?.httpHeaders;
  const delaySeconds = selectedAudio?.delaySeconds ?? 0;
  const cacheKey = input.remote
    ? subtitleAudioCacheKey(
        selectedAudio
          ? JSON.stringify([input.mediaPath, mediaPath, delaySeconds])
          : input.mediaPath,
        audioStreamIndex,
        httpHeaders ?? input.remote.httpHeaders,
      )
    : null;
  if (input.remote?.sessionDirectory && cacheKey) {
    const cached = await readSubtitleAudioCache({
      directory: input.remote.sessionDirectory,
      key: cacheKey,
      destination: wavPath,
    });
    input.signal?.throwIfAborted();
    if (cached) {
      input.onProgress?.({
        stage: 'extract',
        percent: 100,
        message: 'Reusing downloaded audio...',
      });
      return { wavPath, offset: cached.offset };
    }
  }
  input.onProgress?.({ stage: 'extract', message: 'Inspecting audio tracks...' });
  const audio = await probeSubtitleGenerationAudio({
    command: input.ffprobe,
    mediaPath,
    audioStreamIndex,
    httpHeaders,
    signal: input.signal,
  });
  const isRemote = Boolean(input.remote);
  if (isRemote && (!audio.duration || !Number.isFinite(audio.duration) || audio.duration <= 0))
    throw new Error(
      'Stream generation requires a finite episode duration. Live streams are not supported.',
    );
  const original = {
    url: mediaPath,
    probe: audio,
    httpHeaders: httpHeaders ?? { headers: {}, userAgent: null },
  };
  const source =
    input.remote && !selectedAudio
      ? await selectSubtitleGenerationAlternative({
          original,
          alternatives: input.remote.alternatives ?? [],
          ffprobe: input.ffprobe,
          ffmpeg: input.ffmpeg,
          directory: input.directory,
          onProgress: input.onProgress,
          signal: input.signal,
        })
      : { ...original, label: selectedAudio ? 'selected audio stream' : 'current audio track' };
  const httpArgs = isRemote ? subtitleGenerationHttpArgs(source.httpHeaders) : [];
  const extractionMessage =
    source.url === input.mediaPath
      ? 'Extracting audio...'
      : `Extracting audio from ${source.label}...`;
  const stagedHls =
    source.probe.hls && input.remote
      ? await downloadSubtitleGenerationHls({
          mediaPath: source.url,
          directory: input.directory,
          httpHeaders: source.httpHeaders,
          signal: input.signal,
          onProgress: (percent) =>
            input.onProgress?.({
              stage: 'extract',
              percent: Math.floor(percent * 0.9),
              message:
                source.url === input.mediaPath
                  ? 'Downloading episode segments...'
                  : `Downloading ${source.label} segments...`,
            }),
        })
      : null;
  input.onProgress?.({
    stage: 'extract',
    percent: stagedHls ? 90 : 0,
    message: extractionMessage,
  });
  await runSubtitleGenerationProcess({
    command: input.ffmpeg,
    args: [
      '-nostdin',
      '-hide_banner',
      '-loglevel',
      'error',
      ...(stagedHls ? ['-protocol_whitelist', 'file'] : httpArgs),
      '-i',
      stagedHls ?? source.url,
      '-map',
      `0:${source.probe.index}`,
      ...(isRemote ? ['-t', String(audio.duration)] : []),
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
      if (match && source.probe.duration && source.probe.duration > 0) {
        input.onProgress?.({
          stage: 'extract',
          percent: Math.min(
            100,
            Math.floor(
              (stagedHls ? 90 : 0) +
                (Number(match[1]) / 10000 / source.probe.duration) * (stagedHls ? 0.1 : 1),
            ),
          ),
          message: extractionMessage,
        });
      }
    },
  });
  if (stagedHls) await rm(path.dirname(stagedHls), { recursive: true, force: true });

  input.signal?.throwIfAborted();
  if (input.remote?.sessionDirectory && cacheKey) {
    await writeSubtitleAudioCache({
      directory: input.remote.sessionDirectory,
      key: cacheKey,
      wavPath,
      offset: audio.offset + delaySeconds,
    }).catch(() => undefined);
  }
  return { wavPath, offset: audio.offset + delaySeconds };
}
