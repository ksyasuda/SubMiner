import * as childProcess from 'node:child_process';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import type { SampledAudioWindow } from './alass-align';
import {
  decodeSubtitleBytes,
  parseSubtitleFileContent,
  type ParsedSubtitleFile,
  type SubtitleTimeSpan,
} from './subtitle-lines';
import { extensionToFfmpegMuxer, fileExists, resolveCommandInvocation, runCommand } from './utils';

const PCM_VAD_SAMPLE_RATE = 8000;
const PCM_VAD_FILTER = 'highpass=f=200,lowpass=f=3400';
const AUDIO_EXTRACT_TIMEOUT_MS = 180_000;
const SAMPLED_WINDOW_TIMEOUT_MS = 30_000;
const SAMPLED_WINDOW_DURATION_SEC = 60;
const SAMPLED_MIN_MEDIA_DURATION_SEC = 300;
const SAMPLED_WINDOW_FRACTIONS = [0.12, 0.35, 0.6, 0.82] as const;
const MAX_CACHED_AUDIO_SPEECH_ENTRIES = 3;

export type AudioSpeechSource = 'silero-vad' | 'builtin-vad';

export interface AudioSpeechResult {
  spans: SubtitleTimeSpan[];
  durationSec: number;
  calibration: 'clean' | 'noisy';
  source: AudioSpeechSource;
}

export interface ExtractAudioSpeechOptions {
  ffmpegPath: string;
  videoPath: string;
  audioStreamIndex: number | null;
  vadExecutablePath?: string | null;
  vadModelPath?: string | null;
  signal?: AbortSignal;
}

export interface EmbeddedTrackSpec {
  ffIndex: number;
  extension: string;
}

export interface SinglePassRipMediaOptions extends ExtractAudioSpeechOptions {
  embeddedTracks: readonly EmbeddedTrackSpec[];
  includeAudio: boolean;
}

export interface SinglePassRipMediaResult {
  audioSpeech: AudioSpeechResult | null;
  embeddedSubtitles: Map<number, ParsedSubtitleFile>;
}

const audioSpeechCache = new Map<string, AudioSpeechResult>();
const sampledAudioCache = new Map<string, SampledAudioWindow[]>();
const embeddedSubtitlesCache = new Map<string, Map<number, ParsedSubtitleFile>>();

export function clearAudioSpeechCache(): void {
  audioSpeechCache.clear();
  sampledAudioCache.clear();
  embeddedSubtitlesCache.clear();
}

function buildFileSignatureKey(videoPath: string): string | null {
  try {
    const stat = fs.statSync(videoPath);
    if (!stat.isFile()) return null;
    return `${videoPath}\0${stat.size}\0${stat.mtimeMs}`;
  } catch {
    return null;
  }
}

function buildVadConfigKey(
  vadExecutablePath?: string | null,
  vadModelPath?: string | null,
): string {
  const exec = vadExecutablePath?.trim() ?? '';
  const model = vadModelPath?.trim() ?? '';
  if (exec.length > 0 && model.length > 0 && fileExists(model)) {
    return `silero:${exec}\0${model}`;
  }
  return 'builtin';
}

function buildAudioCacheKey(
  videoPath: string,
  audioStreamIndex: number | null,
  vadConfigKey = 'builtin',
): string | null {
  const sig = buildFileSignatureKey(videoPath);
  if (!sig) return null;
  return `${sig}\0${audioStreamIndex ?? 'default'}\0${vadConfigKey}`;
}

export function getCachedAudioSpeechSpans(
  videoPath: string,
  audioStreamIndex: number | null,
  vadConfig?: { vadExecutablePath?: string | null; vadModelPath?: string | null } | null,
): AudioSpeechResult | null {
  const vadKey = buildVadConfigKey(vadConfig?.vadExecutablePath, vadConfig?.vadModelPath);
  const key = buildAudioCacheKey(videoPath, audioStreamIndex, vadKey);
  return key ? (audioSpeechCache.get(key) ?? null) : null;
}

function rememberLru<T>(cache: Map<string, T>, key: string | null, value: T): void {
  if (!key) return;
  if (cache.has(key)) {
    cache.delete(key);
  }
  cache.set(key, value);
  while (cache.size > MAX_CACHED_AUDIO_SPEECH_ENTRIES) {
    const oldestKey = cache.keys().next().value;
    if (oldestKey === undefined) break;
    cache.delete(oldestKey);
  }
}

function rememberAudioSpeech(key: string | null, result: AudioSpeechResult): void {
  rememberLru(audioSpeechCache, key, result);
}

export function getMediaFileSignature(videoPath: string): string | null {
  return buildFileSignatureKey(videoPath);
}

export function getCachedEmbeddedSubtitles(
  videoPath: string,
): Map<number, ParsedSubtitleFile> | null {
  const fileKey = buildFileSignatureKey(videoPath);
  if (!fileKey) return null;
  const cached = embeddedSubtitlesCache.get(fileKey);
  return cached && cached.size > 0 ? new Map(cached) : null;
}

export function cacheEmbeddedSubtitleTrack(
  videoPath: string,
  ffIndex: number,
  parsed: ParsedSubtitleFile,
  expectedFileSignature?: string | null,
): void {
  if (parsed.events.length === 0) return;
  const fileKey = buildFileSignatureKey(videoPath);
  if (!fileKey) return;
  if (expectedFileSignature !== undefined && fileKey !== expectedFileSignature) {
    return;
  }
  const existing = embeddedSubtitlesCache.get(fileKey);
  const updated = new Map<number, ParsedSubtitleFile>(existing);
  updated.set(ffIndex, parsed);
  rememberLru(embeddedSubtitlesCache, fileKey, updated);
}

/**
 * Pure-TypeScript voice-band speech activity detector operating on 16-bit mono
 * PCM (expected 8 kHz, voice band-passed 200 to 3400 Hz). Uses a global dynamic-range
 * histogram with hysteresis, zero-crossing gating, hangover, and short-gap bridging.
 */
export function detectSpeechSpansFromPcm(
  samples: Int16Array,
  sampleRate = PCM_VAD_SAMPLE_RATE,
): { spans: SubtitleTimeSpan[]; durationSec: number } {
  const durationSec = sampleRate > 0 ? samples.length / sampleRate : 0;
  const frameSamples = Math.max(1, Math.round(sampleRate * 0.01)); // 10ms frames
  const frameCount = Math.floor(samples.length / frameSamples);
  if (frameCount === 0) {
    return { spans: [], durationSec };
  }

  const frameDb = new Float32Array(frameCount);
  const frameZcr = new Float32Array(frameCount);
  // Histogram from -90 dB to 0 dB in 0.5 dB bins (181 bins) for O(N) percentiles.
  const HIST_MIN = -90;
  const HIST_BINS = 181;
  const hist = new Int32Array(HIST_BINS);
  let activeFrames = 0;

  for (let f = 0; f < frameCount; f += 1) {
    const base = f * frameSamples;
    let sumSq = 0;
    let crossings = 0;
    let prevSign = samples[base]! >= 0 ? 1 : -1;
    for (let i = 0; i < frameSamples; i += 1) {
      const s = samples[base + i]! / 32768;
      sumSq += s * s;
      const sign = samples[base + i]! >= 0 ? 1 : -1;
      if (i > 0 && sign !== prevSign) crossings += 1;
      prevSign = sign;
    }
    const rms = Math.sqrt(sumSq / frameSamples);
    const db = rms > 1e-6 ? 20 * Math.log10(rms) : -100;
    frameDb[f] = db;
    frameZcr[f] = crossings / frameSamples;

    if (db > -88) {
      const bin = Math.max(0, Math.min(HIST_BINS - 1, Math.round((db - HIST_MIN) * 2)));
      hist[bin] = hist[bin]! + 1;
      activeFrames += 1;
    }
  }

  if (activeFrames < 10) {
    return { spans: [], durationSec };
  }

  const percentileDb = (fraction: number): number => {
    const target = Math.max(1, Math.floor(activeFrames * fraction));
    let acc = 0;
    for (let b = 0; b < HIST_BINS; b += 1) {
      acc += hist[b]!;
      if (acc >= target) {
        return HIST_MIN + b * 0.5;
      }
    }
    return 0;
  };

  const p20 = percentileDb(0.2);
  const p85 = percentileDb(0.85);
  const dynamicRange = Math.max(0, p85 - p20);

  // Set onset threshold cleanly between background noise floor and speech peak.
  const onsetDb =
    dynamicRange >= 8 ? Math.max(p20 + dynamicRange * 0.45, p20 + 5.5) : Math.max(p85 - 5.0, -52);
  const releaseDb = onsetDb - Math.min(4.0, Math.max(2.0, dynamicRange * 0.15));

  // Hysteresis + 120ms hangover.
  const HANGOVER_FRAMES = 12;
  const MIN_SPEECH_FRAMES = 8; // 80ms
  const BRIDGE_GAP_FRAMES = 12; // 120ms
  const voiced = new Uint8Array(frameCount);

  let inSpeech = false;
  let hangover = 0;
  for (let f = 0; f < frameCount; f += 1) {
    const db = frameDb[f]!;
    const zcr = frameZcr[f]!;
    const speechLikeZcr = zcr >= 0.015 && zcr <= 0.65;
    if (!inSpeech) {
      if (db >= onsetDb && speechLikeZcr) {
        inSpeech = true;
        hangover = HANGOVER_FRAMES;
        voiced[f] = 1;
      }
    } else if (db >= releaseDb && speechLikeZcr) {
      hangover = HANGOVER_FRAMES;
      voiced[f] = 1;
    } else if (hangover > 0) {
      hangover -= 1;
      voiced[f] = 1;
    } else {
      inSpeech = false;
    }
  }

  // Extract raw spans and bridge tiny gaps.
  const rawSpans: SubtitleTimeSpan[] = [];
  let startFrame = -1;
  for (let f = 0; f < frameCount; f += 1) {
    if (voiced[f] === 1) {
      if (startFrame === -1) startFrame = f;
    } else if (startFrame !== -1) {
      if (f - startFrame >= MIN_SPEECH_FRAMES) {
        rawSpans.push({ start: startFrame * 0.01, end: f * 0.01 });
      }
      startFrame = -1;
    }
  }
  if (startFrame !== -1 && frameCount - startFrame >= MIN_SPEECH_FRAMES) {
    rawSpans.push({ start: startFrame * 0.01, end: frameCount * 0.01 });
  }

  const bridged: SubtitleTimeSpan[] = [];
  const maxBridgeSec = BRIDGE_GAP_FRAMES * 0.01;
  for (const span of rawSpans) {
    const prev = bridged.at(-1);
    if (prev && span.start - prev.end <= maxBridgeSec) {
      prev.end = span.end;
    } else {
      bridged.push({ ...span });
    }
  }

  return { spans: bridged, durationSec };
}

/**
 * Parse raw Silero speech segments emitted by `whisper-vad-speech-segments`
 * without grouping into 20-second transcription passages.
 */
export function parseRawSileroVadSegments(output: string): SubtitleTimeSpan[] {
  const spans: SubtitleTimeSpan[] = [];
  for (const line of output.split(/\r?\n/)) {
    const match = /^Speech segment \d+: start = (\d+(?:\.\d+)?), end = (\d+(?:\.\d+)?)$/.exec(
      line.trim(),
    );
    if (!match) continue;
    const start = Number(match[1]) / 100;
    const end = Number(match[2]) / 100;
    if (Number.isFinite(start) && Number.isFinite(end) && end > start) {
      spans.push({ start, end });
    }
  }
  return spans;
}

function streamFfmpegPcm(
  ffmpegPath: string,
  videoPath: string,
  audioStreamIndex: number | null,
  signal?: AbortSignal,
): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const args = [
      '-hide_banner',
      '-nostdin',
      '-loglevel',
      'error',
      '-i',
      videoPath,
      '-vn',
      '-sn',
      '-dn',
    ];
    if (audioStreamIndex !== null && Number.isInteger(audioStreamIndex) && audioStreamIndex >= 0) {
      args.push('-map', `0:${audioStreamIndex}`);
    } else {
      args.push('-map', '0:a:0');
    }
    args.push(
      '-ac',
      '1',
      '-ar',
      String(PCM_VAD_SAMPLE_RATE),
      '-af',
      PCM_VAD_FILTER,
      '-f',
      's16le',
      'pipe:1',
    );

    const invocation = resolveCommandInvocation(ffmpegPath, args);
    const child = childProcess.spawn(invocation.command, invocation.args, {
      stdio: ['ignore', 'pipe', 'pipe'],
      ...(invocation.windowsVerbatimArguments ? { windowsVerbatimArguments: true } : {}),
    });
    const chunks: Buffer[] = [];
    let stderr = '';
    let settled = false;

    const finishError = (err: Error): void => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      signal?.removeEventListener('abort', onAbort);
      reject(err);
    };

    const onAbort = (): void => {
      child.kill('SIGKILL');
      finishError(new Error('Audio extraction aborted'));
    };

    if (signal?.aborted) {
      onAbort();
      return;
    }
    signal?.addEventListener('abort', onAbort, { once: true });

    const timer = setTimeout(() => {
      child.kill('SIGKILL');
      finishError(new Error('Timed out while reading video audio with ffmpeg'));
    }, AUDIO_EXTRACT_TIMEOUT_MS);

    child.stdout.on('data', (chunk: Buffer) => {
      chunks.push(chunk);
    });
    child.stderr.on('data', (chunk: Buffer) => {
      stderr += chunk.toString();
    });
    child.on('error', (error: Error) => {
      finishError(error);
    });
    child.on('close', (code: number | null) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      signal?.removeEventListener('abort', onAbort);
      if (code !== 0) {
        reject(new Error(`ffmpeg audio read failed (code=${code ?? 'n/a'}): ${stderr.trim()}`));
        return;
      }
      resolve(Buffer.concat(chunks));
    });
  });
}

async function trySileroVadSpeech(
  options: ExtractAudioSpeechOptions,
): Promise<AudioSpeechResult | null> {
  const vadExec = options.vadExecutablePath?.trim();
  const vadModel = options.vadModelPath?.trim();
  if (!vadExec || !vadModel || !fileExists(vadModel)) {
    return null;
  }

  let tempDir: string | null = null;
  try {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-subsync-vad-'));
    const wavPath = path.join(tempDir, 'audio.wav');
    const ffmpegArgs = [
      '-hide_banner',
      '-nostdin',
      '-y',
      '-loglevel',
      'error',
      '-i',
      options.videoPath,
      '-vn',
      '-sn',
      '-dn',
    ];
    if (
      options.audioStreamIndex !== null &&
      Number.isInteger(options.audioStreamIndex) &&
      options.audioStreamIndex >= 0
    ) {
      ffmpegArgs.push('-map', `0:${options.audioStreamIndex}`);
    } else {
      ffmpegArgs.push('-map', '0:a:0');
    }
    ffmpegArgs.push('-ac', '1', '-ar', '16000', '-c:a', 'pcm_s16le', wavPath);

    const wavRes = await runCommand(options.ffmpegPath, ffmpegArgs, AUDIO_EXTRACT_TIMEOUT_MS);
    if (!wavRes.ok || !fileExists(wavPath)) return null;

    const wavStats = fs.statSync(wavPath);
    const durationSec = Math.max(0, (wavStats.size - 44) / (16000 * 2));

    const vadRes = await runCommand(
      vadExec,
      [
        '-f',
        wavPath,
        '-vm',
        vadModel,
        '-vt',
        '0.35',
        '--vad-min-speech-duration-ms',
        '120',
        '--vad-min-silence-duration-ms',
        '250',
        '-vp',
        '100',
        '-np',
      ],
      AUDIO_EXTRACT_TIMEOUT_MS,
    );
    if (!vadRes.ok) return null;

    const spans = parseRawSileroVadSegments(`${vadRes.stdout}\n${vadRes.stderr}`);
    if (spans.length === 0) return null;

    return {
      spans,
      durationSec: Math.max(durationSec, spans[spans.length - 1]!.end),
      calibration: 'clean',
      source: 'silero-vad',
    };
  } catch {
    return null;
  } finally {
    if (tempDir) {
      try {
        fs.rmSync(tempDir, { recursive: true, force: true });
      } catch {
        // Best-effort cleanup.
      }
    }
  }
}

export async function extractAudioSpeechSpans(
  options: ExtractAudioSpeechOptions,
): Promise<AudioSpeechResult> {
  const vadKey = buildVadConfigKey(options.vadExecutablePath, options.vadModelPath);
  const cacheKey = buildAudioCacheKey(options.videoPath, options.audioStreamIndex, vadKey);
  if (cacheKey) {
    const cached = audioSpeechCache.get(cacheKey);
    if (cached) return cached;
  }

  // Tier 1: Prefer Silero VAD if whisper-vad-speech-segments + model are available.
  const silero = await trySileroVadSpeech(options);
  if (silero) {
    rememberAudioSpeech(cacheKey, silero);
    return silero;
  }

  // Tier 2: Built-in voice-band PCM speech detector via ffmpeg (zero extra dependencies).
  const pcmBuffer = await streamFfmpegPcm(
    options.ffmpegPath,
    options.videoPath,
    options.audioStreamIndex,
    options.signal,
  );
  const sampleCount = Math.floor(pcmBuffer.byteLength / 2);
  const samples = new Int16Array(pcmBuffer.buffer, pcmBuffer.byteOffset, sampleCount);
  const { spans, durationSec } = detectSpeechSpansFromPcm(samples, PCM_VAD_SAMPLE_RATE);

  const result: AudioSpeechResult = {
    spans,
    durationSec,
    calibration: 'noisy',
    source: 'builtin-vad',
  };
  rememberAudioSpeech(cacheKey, result);
  return result;
}

function spawnFfmpegWithStdout(
  ffmpegPath: string,
  args: string[],
  timeoutMs: number,
  signal?: AbortSignal,
): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const invocation = resolveCommandInvocation(ffmpegPath, args);
    const child = childProcess.spawn(invocation.command, invocation.args, {
      stdio: ['ignore', 'pipe', 'pipe'],
      ...(invocation.windowsVerbatimArguments ? { windowsVerbatimArguments: true } : {}),
    });
    const chunks: Buffer[] = [];
    let stderr = '';
    let settled = false;

    const finishError = (err: Error): void => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      signal?.removeEventListener('abort', onAbort);
      reject(err);
    };

    const onAbort = (): void => {
      child.kill('SIGKILL');
      finishError(new Error('ffmpeg operation aborted'));
    };

    if (signal?.aborted) {
      onAbort();
      return;
    }
    signal?.addEventListener('abort', onAbort, { once: true });

    const timer = setTimeout(() => {
      child.kill('SIGKILL');
      finishError(new Error('ffmpeg operation timed out'));
    }, timeoutMs);

    child.stdout.on('data', (chunk: Buffer) => {
      chunks.push(chunk);
    });
    child.stderr.on('data', (chunk: Buffer) => {
      stderr += chunk.toString();
    });
    child.on('error', (error: Error) => {
      finishError(error);
    });
    child.on('close', (code: number | null) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      signal?.removeEventListener('abort', onAbort);
      if (code !== 0) {
        reject(new Error(`ffmpeg failed (code=${code ?? 'n/a'}): ${stderr.trim()}`));
        return;
      }
      resolve(Buffer.concat(chunks));
    });
  });
}

function audioMapArg(audioStreamIndex: number | null): string {
  return audioStreamIndex !== null && Number.isInteger(audioStreamIndex) && audioStreamIndex >= 0
    ? `0:${audioStreamIndex}`
    : '0:a:0';
}

/**
 * Sample 4 seeked 60-second audio windows across a media file in parallel using
 * fast input-side `-ss` seeks. Used to verify external reference subtitles in
 * ~0.2-0.5s without decoding the entire episode or movie.
 */
export async function extractSampledAudioWindows(options: {
  ffmpegPath: string;
  videoPath: string;
  audioStreamIndex: number | null;
  durationSec: number;
  signal?: AbortSignal;
}): Promise<SampledAudioWindow[] | null> {
  if (
    !Number.isFinite(options.durationSec) ||
    options.durationSec < SAMPLED_MIN_MEDIA_DURATION_SEC
  ) {
    return null;
  }

  const audioKey = buildAudioCacheKey(options.videoPath, options.audioStreamIndex);
  const sampleKey = audioKey ? `${audioKey}\0${Math.round(options.durationSec)}` : null;
  if (sampleKey) {
    const cached = sampledAudioCache.get(sampleKey);
    if (cached) return cached;
  }

  const maxStart = Math.max(0, options.durationSec - SAMPLED_WINDOW_DURATION_SEC - 5);
  const starts: number[] = [];
  for (const frac of SAMPLED_WINDOW_FRACTIONS) {
    const startSec = Math.round(Math.min(maxStart, Math.max(0, options.durationSec * frac)));
    if (starts.every((prev) => Math.abs(prev - startSec) >= SAMPLED_WINDOW_DURATION_SEC)) {
      starts.push(startSec);
    }
  }
  if (starts.length < 2) return null;

  const mapSpec = audioMapArg(options.audioStreamIndex);
  const maxBytes = SAMPLED_WINDOW_DURATION_SEC * PCM_VAD_SAMPLE_RATE * 2;

  const settled = await Promise.all(
    starts.map(async (startSec): Promise<SampledAudioWindow | null> => {
      try {
        const pcmBuffer = await spawnFfmpegWithStdout(
          options.ffmpegPath,
          [
            '-hide_banner',
            '-nostdin',
            '-loglevel',
            'error',
            '-ss',
            String(startSec),
            '-t',
            String(SAMPLED_WINDOW_DURATION_SEC),
            '-i',
            options.videoPath,
            '-vn',
            '-sn',
            '-dn',
            '-map',
            mapSpec,
            '-ac',
            '1',
            '-ar',
            String(PCM_VAD_SAMPLE_RATE),
            '-af',
            PCM_VAD_FILTER,
            '-f',
            's16le',
            'pipe:1',
          ],
          SAMPLED_WINDOW_TIMEOUT_MS,
          options.signal,
        );
        const clampedLen = Math.min(pcmBuffer.byteLength, maxBytes);
        const sampleCount = Math.floor(clampedLen / 2);
        if (sampleCount < PCM_VAD_SAMPLE_RATE * 10) return null;
        const samples = new Int16Array(pcmBuffer.buffer, pcmBuffer.byteOffset, sampleCount);
        const { spans, durationSec } = detectSpeechSpansFromPcm(samples, PCM_VAD_SAMPLE_RATE);
        if (spans.length === 0) return null;
        return {
          startSec,
          durationSec: Math.min(SAMPLED_WINDOW_DURATION_SEC, durationSec),
          speechSpans: spans,
        };
      } catch {
        return null;
      }
    }),
  );

  const valid = settled.filter((w): w is SampledAudioWindow => w !== null);
  if (valid.length < 2) return null;

  rememberLru(sampledAudioCache, sampleKey, valid);
  return valid;
}

/**
 * Extract multiple embedded text subtitle tracks AND (optionally) audio speech
 * in ONE single `ffmpeg` demux pass, avoiding N+1 full-file reads on multi-track
 * MKVs and caching results for subsequent modal retries.
 * Returns `null` if the multi-output ffmpeg pass fails so callers can fall back.
 */
export async function extractSinglePassRipMedia(
  options: SinglePassRipMediaOptions,
): Promise<SinglePassRipMediaResult | null> {
  const fileKey = buildFileSignatureKey(options.videoPath);
  const vadKey = buildVadConfigKey(options.vadExecutablePath, options.vadModelPath);
  const audioKey = buildAudioCacheKey(options.videoPath, options.audioStreamIndex, vadKey);

  const cachedSubs =
    (fileKey ? embeddedSubtitlesCache.get(fileKey) : undefined) ??
    new Map<number, ParsedSubtitleFile>();
  const cachedAudio = audioKey ? (audioSpeechCache.get(audioKey) ?? null) : null;

  const missingTracks = options.embeddedTracks.filter((t) => !cachedSubs.has(t.ffIndex));
  const needAudio = options.includeAudio && cachedAudio === null;

  if (missingTracks.length === 0 && !needAudio) {
    return {
      audioSpeech: options.includeAudio ? cachedAudio : null,
      embeddedSubtitles: new Map(cachedSubs),
    };
  }

  const vadExec = options.vadExecutablePath?.trim() ?? '';
  const vadModel = options.vadModelPath?.trim() ?? '';
  const useSilero = needAudio && vadExec.length > 0 && vadModel.length > 0 && fileExists(vadModel);

  let tempDir: string | null = null;

  try {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-subsync-rip-'));
    const wavPath = path.join(tempDir, 'audio.wav');
    const args: string[] = [
      '-hide_banner',
      '-nostdin',
      '-y',
      '-loglevel',
      'error',
      '-i',
      options.videoPath,
    ];

    const trackOutputs: Array<{ ffIndex: number; outPath: string }> = [];
    for (const track of missingTracks) {
      const outPath = path.join(tempDir, `track_${track.ffIndex}.${track.extension}`);
      args.push(
        '-vn',
        '-an',
        '-map',
        `0:${track.ffIndex}`,
        '-f',
        extensionToFfmpegMuxer(track.extension),
        outPath,
      );
      trackOutputs.push({ ffIndex: track.ffIndex, outPath });
    }

    if (needAudio) {
      const mapSpec = audioMapArg(options.audioStreamIndex);
      if (useSilero) {
        args.push(
          '-vn',
          '-sn',
          '-dn',
          '-map',
          mapSpec,
          '-ac',
          '1',
          '-ar',
          '16000',
          '-c:a',
          'pcm_s16le',
          wavPath,
        );
      } else {
        args.push(
          '-vn',
          '-sn',
          '-dn',
          '-map',
          mapSpec,
          '-ac',
          '1',
          '-ar',
          String(PCM_VAD_SAMPLE_RATE),
          '-af',
          PCM_VAD_FILTER,
          '-f',
          's16le',
          'pipe:1',
        );
      }
    }

    const stdoutBuf = await spawnFfmpegWithStdout(
      options.ffmpegPath,
      args,
      AUDIO_EXTRACT_TIMEOUT_MS,
      options.signal,
    );

    const updatedSubs = new Map<number, ParsedSubtitleFile>(cachedSubs);
    for (const { ffIndex, outPath } of trackOutputs) {
      if (!fileExists(outPath)) continue;
      try {
        const bytes = fs.readFileSync(outPath);
        const parsed = parseSubtitleFileContent(decodeSubtitleBytes(bytes), outPath);
        updatedSubs.set(ffIndex, parsed);
      } catch {
        // Ignore individual unparseable track.
      }
    }
    if (updatedSubs.size > 0) {
      rememberLru(embeddedSubtitlesCache, fileKey, updatedSubs);
    }

    let audioSpeech: AudioSpeechResult | null = options.includeAudio ? cachedAudio : null;
    if (needAudio) {
      if (useSilero && fileExists(wavPath)) {
        const wavStats = fs.statSync(wavPath);
        const durationSec = Math.max(0, (wavStats.size - 44) / (16000 * 2));
        const vadRes = await runCommand(
          vadExec,
          [
            '-f',
            wavPath,
            '-vm',
            vadModel,
            '-vt',
            '0.35',
            '--vad-min-speech-duration-ms',
            '120',
            '--vad-min-silence-duration-ms',
            '250',
            '-vp',
            '100',
            '-np',
          ],
          AUDIO_EXTRACT_TIMEOUT_MS,
        );
        if (vadRes.ok) {
          const spans = parseRawSileroVadSegments(`${vadRes.stdout}\n${vadRes.stderr}`);
          if (spans.length > 0) {
            audioSpeech = {
              spans,
              durationSec: Math.max(durationSec, spans[spans.length - 1]!.end),
              calibration: 'clean',
              source: 'silero-vad',
            };
          }
        }
        // If Silero returned no spans, fall back cleanly to built-in PCM VAD.
        if (!audioSpeech) {
          try {
            audioSpeech = await extractAudioSpeechSpans({
              ffmpegPath: options.ffmpegPath,
              videoPath: options.videoPath,
              audioStreamIndex: options.audioStreamIndex,
              signal: options.signal,
            });
          } catch {
            audioSpeech = null;
          }
        }
      } else if (!useSilero) {
        const sampleCount = Math.floor(stdoutBuf.byteLength / 2);
        if (sampleCount > 0) {
          const samples = new Int16Array(stdoutBuf.buffer, stdoutBuf.byteOffset, sampleCount);
          const { spans, durationSec } = detectSpeechSpansFromPcm(samples, PCM_VAD_SAMPLE_RATE);
          if (spans.length > 0) {
            audioSpeech = {
              spans,
              durationSec,
              calibration: 'noisy',
              source: 'builtin-vad',
            };
          }
        }
      }

      if (audioSpeech) {
        rememberAudioSpeech(audioKey, audioSpeech);
      }
    }

    return {
      audioSpeech,
      embeddedSubtitles: updatedSubs,
    };
  } catch {
    return null;
  } finally {
    if (tempDir) {
      try {
        fs.rmSync(tempDir, { recursive: true, force: true });
      } catch {
        // Best-effort cleanup.
      }
    }
  }
}
