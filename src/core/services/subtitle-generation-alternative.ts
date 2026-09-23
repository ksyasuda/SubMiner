import { mkdtemp, readFile, rm } from 'node:fs/promises';
import path from 'node:path';
import { runSubtitleGenerationProcess } from './subtitle-generation-process';
import {
  probeSubtitleGenerationAudio,
  type SubtitleGenerationAudioProbe,
} from './subtitle-generation-probe';
import {
  subtitleGenerationHttpArgs,
  type SubtitleGenerationAlternative,
} from './subtitle-generation-source';
import type { ResolvedMpvHttpHeaders } from './mpv-http-headers';
import type { SubtitleGenerationProgress } from '../../shared/subtitle-generation';

import { downloadSubtitleGenerationHlsWindow } from './subtitle-generation-hls';
import { createLogger } from '../../logger';

const logger = createLogger('subtitle-generation:sources');
export interface SubtitleSourceDiagnostic {
  candidate: number;
  kind: SubtitleGenerationAlternative['kind'];
  phase: 'probe' | 'original-sample' | 'alternative-sample' | 'comparison';
  result: 'retry' | 'rejected' | 'selected';
  reason:
    | 'empty-sample'
    | 'short-sample'
    | 'duration-mismatch'
    | 'offset-mismatch'
    | 'audio-mismatch'
    | 'timeout'
    | 'process-error'
    | 'matched';
  elapsedMs: number;
}

class InvalidSample extends Error {
  constructor(readonly reason: 'empty-sample' | 'short-sample') {
    super(reason);
  }
}

interface AudioSource {
  url: string;
  httpHeaders: ResolvedMpvHttpHeaders;
  probe: SubtitleGenerationAudioProbe;
}

// A duration match alone accepts dubs and differently edited releases. Compare
// decoded audio at the same times too; uncertain matches keep the playback source.
function matchingSamples(a: Buffer, b: Buffer): boolean {
  // FFmpeg resampling can round a two-second extract by a few trailing samples.
  if (
    Math.min(a.length, b.length) < 16000 ||
    Math.abs(a.length - b.length) > 80 ||
    a.length % 2 !== 0 ||
    b.length % 2 !== 0
  )
    return false;
  // Container seek precision and resampling can differ by fractional audio samples.
  // This tolerance cannot hide a materially shifted subtitle timeline.
  const count = Math.min(a.length, b.length) / 2;
  for (let lag = -40; lag <= 40; lag += 0.25) {
    let sumA = 0,
      sumB = 0,
      aa = 0,
      bb = 0,
      ab = 0,
      samples = 0;
    // Interpolation needs the next sample even at the maximum positive lag.
    for (let i = 40; i < count - 41; i += 2) {
      const index = Math.floor(i + lag);
      const fraction = i + lag - index;
      const x = a.readInt16LE(i * 2);
      const y =
        b.readInt16LE(index * 2) * (1 - fraction) + b.readInt16LE((index + 1) * 2) * fraction;
      sumA += x;
      sumB += y;
      aa += x * x;
      bb += y * y;
      ab += x * y;
      samples += 1;
    }
    const varianceA = aa - (sumA * sumA) / samples;
    const varianceB = bb - (sumB * sumB) / samples;
    if (varianceA / samples < 100 || varianceB / samples < 100) return false;
    if ((ab - (sumA * sumB) / samples) / Math.sqrt(varianceA * varianceB) >= 0.98) return true;
  }
  return false;
}

export async function selectSubtitleGenerationAlternative(input: {
  original: AudioSource;
  alternatives: readonly SubtitleGenerationAlternative[];
  ffprobe: string;
  ffmpeg: string;
  directory: string;
  signal?: AbortSignal;
  onProgress?: (progress: SubtitleGenerationProgress) => void;
  onDiagnostic?: (event: SubtitleSourceDiagnostic) => void;
}): Promise<AudioSource & { label: string }> {
  const original = { ...input.original, label: 'current audio track' };
  const duration = original.probe.duration;
  if (!duration || duration < 12 || !input.alternatives.length) return original;
  const started = Date.now();
  const budget = AbortSignal.timeout(20_000);
  const signal = input.signal ? AbortSignal.any([input.signal, budget]) : budget;
  const times = [0.15, 0.5, 0.85].map((fraction) => Math.min(duration - 3, duration * fraction));
  const originals = new Map<number, Buffer>();
  const sample = async (source: AudioSource, seconds: number) => {
    const directory = await mkdtemp(path.join(input.directory, 'comparison-'));
    const output = path.join(directory, 'sample.pcm');
    const timeout = AbortSignal.timeout(15_000);
    try {
      const sampleSignal = AbortSignal.any([signal, timeout]);
      const staged = source.probe.hls
        ? await downloadSubtitleGenerationHlsWindow({
            mediaPath: source.url,
            directory,
            httpHeaders: source.httpHeaders,
            signal: sampleSignal,
            window: { startSeconds: Math.max(0, seconds - 0.1), durationSeconds: 2.2 },
          })
        : null;
      const localProbe = staged
        ? await probeSubtitleGenerationAudio({
            command: input.ffprobe,
            mediaPath: staged.playlistPath,
            audioStreamIndex: source.probe.index,
            signal: sampleSignal,
          })
        : null;
      // Segment timestamps, not rounded EXTINF sums, locate the exact playback time.
      const seek = localProbe ? seconds + source.probe.startTime - localProbe.startTime : seconds;
      if (seek < 0) throw new InvalidSample('short-sample');
      await runSubtitleGenerationProcess({
        command: input.ffmpeg,
        args: [
          '-nostdin',
          '-v',
          'error',
          ...(staged
            ? ['-protocol_whitelist', 'file']
            : subtitleGenerationHttpArgs(source.httpHeaders)),
          ...(staged ? [] : ['-ss', String(seconds)]),
          '-i',
          staged?.playlistPath ?? source.url,
          ...(staged ? ['-ss', String(seek)] : []),
          '-map',
          `0:${source.probe.index}`,
          '-t',
          '2',
          '-vn',
          '-ac',
          '1',
          '-ar',
          '8000',
          '-c:a',
          'pcm_s16le',
          '-f',
          's16le',
          output,
        ],
        signal: sampleSignal,
      });
      const bytes = await readFile(output);
      if (bytes.length === 0) throw new InvalidSample('empty-sample');
      if (bytes.length < 16000 || bytes.length % 2 !== 0) throw new InvalidSample('short-sample');
      return bytes;
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  };
  for (const [candidateIndex, alternative] of input.alternatives.slice(0, 4).entries()) {
    input.signal?.throwIfAborted();
    if (budget.aborted) break;
    if (!/^https?:\/\//i.test(alternative.url) || alternative.url === original.url) continue;
    input.onProgress?.({
      stage: 'extract',
      message: `Checking ${alternative.label} audio and timing...`,
    });
    let phase: SubtitleSourceDiagnostic['phase'] = 'probe';
    const report = (
      result: SubtitleSourceDiagnostic['result'],
      reason: SubtitleSourceDiagnostic['reason'],
    ) => {
      const event = {
        candidate: candidateIndex + 1,
        kind: alternative.kind,
        phase,
        result,
        reason,
        elapsedMs: Date.now() - started,
      };
      input.onDiagnostic?.(event);
      // Never log stream URLs, headers or raw FFmpeg errors containing credentials.
      if (result === 'selected') logger.info('Alternative source selected', event);
      else logger.warn('Alternative source check', event);
    };
    try {
      const timeout = AbortSignal.timeout(15_000);
      const probe = await probeSubtitleGenerationAudio({
        command: input.ffprobe,
        mediaPath: alternative.url,
        httpHeaders: alternative.httpHeaders,
        signal: AbortSignal.any([signal, timeout]),
      });
      if (!probe.duration || Math.abs(probe.duration - duration) > 0.25) {
        report('rejected', 'duration-mismatch');
        continue;
      }
      if (Math.abs(probe.offset - original.probe.offset) > 0.05) {
        report('rejected', 'offset-mismatch');
        continue;
      }
      const candidate = {
        url: alternative.url,
        httpHeaders: alternative.httpHeaders,
        probe,
        label: alternative.label,
      };
      let matches = true;
      for (const seconds of times) {
        let baseline = originals.get(seconds);
        if (!baseline) {
          phase = 'original-sample';
          try {
            baseline = await sample(original, seconds);
          } catch (error) {
            if (!(error instanceof InvalidSample) || signal.aborted) throw error;
            report('retry', error.reason);
            baseline = await sample(original, seconds);
          }
          // Invalid samples never enter the reusable baseline cache.
          originals.set(seconds, baseline);
        }
        phase = 'alternative-sample';
        const comparison = await sample(candidate, seconds);
        phase = 'comparison';
        if (!matchingSamples(baseline, comparison)) {
          report('rejected', 'audio-mismatch');
          matches = false;
          break;
        }
      }
      if (matches) {
        report('selected', 'matched');
        return candidate;
      }
    } catch (error) {
      input.signal?.throwIfAborted();
      report(
        'rejected',
        error instanceof InvalidSample
          ? error.reason
          : signal.aborted || (error instanceof Error && /cancelled|timed out/i.test(error.message))
            ? 'timeout'
            : 'process-error',
      );
    }
  }
  return original;
}
