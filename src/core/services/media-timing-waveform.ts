import { spawn } from 'node:child_process';
import { normalizeMediaInput, type MediaInput } from '../../media-input';

const WAVEFORM_SAMPLE_RATE = 8_000;
const WAVEFORM_POINT_COUNT = 480;
const WAVEFORM_TIMEOUT_MS = 15_000;
const MAX_WAVEFORM_BYTES = 16 * 1024 * 1024;
// Keep the band where speech intelligibility lives; bass, drums, and hum sit below it.
const SPEECH_FILTER = 'highpass=f=250,lowpass=f=3500';
const NOISE_FLOOR_PERCENTILE = 0.2;
const REFERENCE_PERCENTILE = 0.95;
const NOISE_GATE_DB = 3;
const MIN_DISPLAY_RANGE_DB = 12;
const SILENCE_DB = -100;
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

function percentile(sortedValues: number[], fraction: number): number {
  const index = Math.min(sortedValues.length - 1, Math.floor(sortedValues.length * fraction));
  return sortedValues[index] ?? SILENCE_DB;
}

/**
 * Turns mono PCM into 0..1 display heights. Each point is the RMS level of its slice in
 * dB, measured against the clip's own noise floor (a low percentile of the slices), so
 * constant background noise draws flat and sustained speech stands out. Peak sampling
 * would instead follow music transients and lift the floor to nearly speech height.
 */
export function computeWaveformPeaks(pcm: Buffer, pointCount = WAVEFORM_POINT_COUNT): number[] {
  const sampleCount = Math.floor(pcm.byteLength / 2);
  if (sampleCount === 0 || pointCount <= 0) return [];
  const resolvedPointCount = Math.min(pointCount, sampleCount);
  const levelsDb = Array.from({ length: resolvedPointCount }, () => SILENCE_DB);

  for (let point = 0; point < resolvedPointCount; point += 1) {
    const sampleStart = Math.floor((point * sampleCount) / resolvedPointCount);
    const sampleEnd = Math.max(
      sampleStart + 1,
      Math.floor(((point + 1) * sampleCount) / resolvedPointCount),
    );
    let energy = 0;
    for (let sample = sampleStart; sample < sampleEnd; sample += 1) {
      const value = pcm.readInt16LE(sample * 2) / 32_768;
      energy += value * value;
    }
    const rms = Math.sqrt(energy / (sampleEnd - sampleStart));
    levelsDb[point] = rms > 0 ? Math.max(SILENCE_DB, 20 * Math.log10(rms)) : SILENCE_DB;
  }

  const sortedLevels = [...levelsDb].sort((left, right) => left - right);
  const floorDb = percentile(sortedLevels, NOISE_FLOOR_PERCENTILE) + NOISE_GATE_DB;
  const referenceDb = Math.max(
    percentile(sortedLevels, REFERENCE_PERCENTILE),
    floorDb + MIN_DISPLAY_RANGE_DB,
  );
  return levelsDb.map(
    (levelDb) =>
      Math.round(Math.min(1, Math.max(0, (levelDb - floorDb) / (referenceDb - floorDb))) * 1_000) /
      1_000,
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
