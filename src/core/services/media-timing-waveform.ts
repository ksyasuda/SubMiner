import { spawn } from 'node:child_process';
import { normalizeMediaInput, type MediaInput } from '../../media-input';

const WAVEFORM_SAMPLE_RATE = 8_000;
const WAVEFORM_POINT_COUNT = 480;
const WAVEFORM_TIMEOUT_MS = 15_000;
const MAX_WAVEFORM_BYTES = 16 * 1024 * 1024;
const SPEECH_FILTER = 'highpass=f=120,lowpass=f=4000';
const CENTER_CHANNEL_FILTER = `pan=mono|c0=FC,${SPEECH_FILTER}`;
const DOWNMIX_FILTER = `aformat=channel_layouts=mono,${SPEECH_FILTER}`;

export interface SpeechWaveformOptions {
  mediaPath: MediaInput;
  startTime: number;
  endTime: number;
  audioStreamIndex?: number;
}

type RunFfmpeg = (args: string[]) => Promise<Buffer>;

export function buildSpeechWaveformArgs(
  options: SpeechWaveformOptions,
  mode: 'center' | 'downmix',
): string[] {
  const duration = options.endTime - options.startTime;
  const input = normalizeMediaInput(options.mediaPath);
  const args = [
    '-hide_banner',
    '-nostdin',
    '-loglevel',
    'error',
    '-ss',
    String(options.startTime),
    ...input.inputArgs,
    '-i',
    input.path,
    '-t',
    String(duration),
  ];
  if (
    options.audioStreamIndex !== undefined &&
    Number.isInteger(options.audioStreamIndex) &&
    options.audioStreamIndex >= 0
  ) {
    args.push('-map', `0:${options.audioStreamIndex}`);
  }
  args.push(
    '-vn',
    '-sn',
    '-dn',
    '-af',
    mode === 'center' ? CENTER_CHANNEL_FILTER : DOWNMIX_FILTER,
    '-ac',
    '1',
    '-ar',
    String(WAVEFORM_SAMPLE_RATE),
    '-f',
    's16le',
    'pipe:1',
  );
  return args;
}

function runFfmpeg(args: string[]): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const child = spawn('ffmpeg', args, { stdio: ['ignore', 'pipe', 'pipe'] });
    const chunks: Buffer[] = [];
    let byteLength = 0;
    let stderr = '';
    let settled = false;
    const timeout = setTimeout(() => {
      if (settled) return;
      settled = true;
      child.kill('SIGKILL');
      reject(new Error(`FFmpeg waveform analysis timed out after ${WAVEFORM_TIMEOUT_MS}ms`));
    }, WAVEFORM_TIMEOUT_MS);

    const settle = (callback: () => void): void => {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      callback();
    };

    child.stdout.on('data', (chunk: Buffer) => {
      if (settled) return;
      byteLength += chunk.byteLength;
      if (byteLength > MAX_WAVEFORM_BYTES) {
        settle(() => {
          child.kill('SIGKILL');
          reject(new Error('The visible waveform range is too large to analyze.'));
        });
        return;
      }
      chunks.push(chunk);
    });
    child.stderr.setEncoding('utf8');
    child.stderr.on('data', (chunk) => {
      if (stderr.length < 4_000) stderr += String(chunk);
    });
    child.once('error', (error) => settle(() => reject(error)));
    child.once('close', (code) => {
      settle(() => {
        if (code === 0) {
          resolve(Buffer.concat(chunks, byteLength));
          return;
        }
        reject(new Error(stderr.trim() || `FFmpeg exited with status ${code ?? 'unknown'}`));
      });
    });
  });
}

export function computeWaveformPeaks(pcm: Buffer, pointCount = WAVEFORM_POINT_COUNT): number[] {
  const sampleCount = Math.floor(pcm.byteLength / 2);
  if (sampleCount === 0 || pointCount <= 0) return [];
  const resolvedPointCount = Math.min(pointCount, sampleCount);
  const peaks = Array.from({ length: resolvedPointCount }, () => 0);

  for (let point = 0; point < resolvedPointCount; point += 1) {
    const sampleStart = Math.floor((point * sampleCount) / resolvedPointCount);
    const sampleEnd = Math.max(
      sampleStart + 1,
      Math.floor(((point + 1) * sampleCount) / resolvedPointCount),
    );
    let peak = 0;
    for (let sample = sampleStart; sample < sampleEnd; sample += 1) {
      peak = Math.max(peak, Math.abs(pcm.readInt16LE(sample * 2)) / 32_768);
    }
    peaks[point] = peak;
  }

  const sortedPeaks = [...peaks].sort((left, right) => left - right);
  const referenceIndex = Math.min(sortedPeaks.length - 1, Math.floor(sortedPeaks.length * 0.95));
  const referencePeak = Math.max(sortedPeaks[referenceIndex] ?? 0, 0.01);
  return peaks.map(
    (peak) => Math.round(Math.sqrt(Math.min(1, peak / referencePeak)) * 1_000) / 1_000,
  );
}

function hasAudibleSamples(pcm: Buffer): boolean {
  for (let offset = 0; offset + 1 < pcm.byteLength; offset += 2) {
    if (Math.abs(pcm.readInt16LE(offset)) >= 164) return true;
  }
  return false;
}

export async function generateSpeechWaveform(
  options: SpeechWaveformOptions,
  execute: RunFfmpeg = runFfmpeg,
): Promise<number[]> {
  try {
    const centerPcm = await execute(buildSpeechWaveformArgs(options, 'center'));
    if (hasAudibleSamples(centerPcm)) return computeWaveformPeaks(centerPcm);
  } catch {
    // Sources without a named center channel can reject the center-only filter.
  }

  const downmixPcm = await execute(buildSpeechWaveformArgs(options, 'downmix'));
  return computeWaveformPeaks(downmixPcm);
}
