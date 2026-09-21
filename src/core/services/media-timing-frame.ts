import { execFile } from 'node:child_process';
import { normalizeMediaInput, type MediaInput } from '../../media-input';

export interface MediaTimingFrameOptions {
  media: MediaInput;
  timestamp: number;
  direction?: -1 | 1;
}

const FRAME_EPSILON = 0.00001;
const PROBE_RADIUS_SECONDS = 2;

function run(file: string, args: string[]): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    execFile(
      file,
      args,
      { encoding: 'buffer', timeout: 30_000, maxBuffer: 8 * 1024 * 1024, windowsHide: true },
      (error, stdout) => {
        // Child-process errors contain the input URL, which may contain authentication tokens.
        if (error)
          reject(
            new Error('Screenshot preview is unavailable. Check FFmpeg and the video source.'),
          );
        else resolve(stdout);
      },
    );
  });
}

/** Uses decoded timestamps, rather than an assumed FPS, including for variable-rate video. */
export function selectMediaTimingFrame(
  times: readonly number[],
  timestamp: number,
  direction?: -1 | 1,
): number | undefined {
  if (direction === -1)
    return [...times].reverse().find((time) => time < timestamp - FRAME_EPSILON);
  if (direction === 1) return times.find((time) => time > timestamp + FRAME_EPSILON);
  return times.find((time) => time >= timestamp - FRAME_EPSILON) ?? times.at(-1);
}

export function createMediaTimingFrameExtractor(execute: typeof run = run) {
  let cached: { key: string; start: number; end: number; times: number[] } | null = null;
  let source: { key: string; offset: number } | null = null;
  let generation = 0;

  async function generate(
    options: MediaTimingFrameOptions,
  ): Promise<{ dataUrl: string; timestamp: number }> {
    const input = normalizeMediaInput(options.media);
    const key = JSON.stringify(options.media);
    const currentGeneration = generation;
    const absolute = typeof options.media !== 'string' && options.media.absoluteTimestamps;
    // -seek_timestamp is an ffmpeg option; ffprobe intervals already use stream timestamps.
    const probeInputArgs = normalizeMediaInput({
      path: input.path,
      ...(typeof options.media !== 'string' ? { inputOptions: options.media.inputOptions } : {}),
    }).inputArgs;
    let offset = source?.key === key ? source.offset : undefined;
    if (offset === undefined) {
      const metadata = JSON.parse(
        (
          await execute('ffprobe', [
            '-v',
            'error',
            ...probeInputArgs,
            '-show_entries',
            'format=start_time',
            '-of',
            'json',
            input.path,
          ])
        ).toString(),
      ) as { format?: { start_time?: string } };
      const start = Number(metadata.format?.start_time ?? 0);
      offset = absolute || !Number.isFinite(start) ? 0 : start;
      if (generation === currentGeneration) source = { key, offset };
    }
    let times: number[];
    if (
      cached?.key === key &&
      options.timestamp > cached.start + 0.5 &&
      options.timestamp < cached.end - 0.5
    ) {
      times = cached.times;
    } else {
      const start = Math.max(0, options.timestamp - PROBE_RADIUS_SECONDS);
      const end = options.timestamp + PROBE_RADIUS_SECONDS;
      const result = JSON.parse(
        (
          await execute('ffprobe', [
            '-v',
            'error',
            ...probeInputArgs,
            '-read_intervals',
            `${start + offset}%${end + offset}`,
            '-select_streams',
            'v:0',
            '-show_entries',
            'frame=best_effort_timestamp_time',
            '-of',
            'json',
            input.path,
          ])
        ).toString(),
      ) as { frames?: { best_effort_timestamp_time?: string }[] };
      times = [
        ...new Set(
          (result.frames ?? [])
            .map((frame) => Number(frame.best_effort_timestamp_time) - offset)
            .filter((time) => Number.isFinite(time) && time >= 0 && time >= start && time <= end),
        ),
      ].sort((a, b) => a - b);
      if (generation === currentGeneration) cached = { key, start, end, times };
    }
    const timestamp = selectMediaTimingFrame(times, options.timestamp, options.direction);
    if (timestamp === undefined) throw new Error('No adjacent video frame is available here.');
    // Round down by one microsecond so decimal timestamp rounding cannot skip the chosen frame.
    const seekTime = Math.max(0, timestamp - 0.000001);
    const image = await execute('ffmpeg', [
      '-hide_banner',
      '-nostdin',
      '-loglevel',
      'error',
      '-ss',
      String(seekTime),
      ...input.inputArgs,
      '-i',
      input.path,
      '-map',
      '0:v:0',
      '-frames:v',
      '1',
      '-an',
      '-sn',
      '-vf',
      'scale=w=640:h=360:force_original_aspect_ratio=decrease',
      '-c:v',
      'mjpeg',
      '-q:v',
      '3',
      '-f',
      'image2pipe',
      'pipe:1',
    ]);
    if (!image.length) throw new Error('No video frame is available here.');
    return { dataUrl: `data:image/jpeg;base64,${image.toString('base64')}`, timestamp: seekTime };
  }

  return {
    generate,
    clear: () => {
      generation += 1;
      cached = null;
      source = null;
    },
  };
}
