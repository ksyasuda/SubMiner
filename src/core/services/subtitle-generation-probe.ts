import { runSubtitleGenerationProcess } from './subtitle-generation-process';
import { subtitleGenerationHttpArgs } from './subtitle-generation-source';
import type { ResolvedMpvHttpHeaders } from './mpv-http-headers';

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
        language: typeof language === 'string' ? language : '',
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
  const hls =
    typeof format === 'object' &&
    format !== null &&
    'format_name' in format &&
    typeof format.format_name === 'string' &&
    format.format_name.split(',').includes('hls');
  return {
    index: selected.index,
    startTime: formatStart,
    language: selected.language,
    // mpv rebases media timestamps to the container start. Extraction rebases the selected audio.
    offset: (selected.start ?? formatStart) - formatStart,
    duration,
    hls,
  };
}

export async function probeSubtitleGenerationAudio(input: {
  command: string;
  mediaPath: string;
  audioStreamIndex?: number;
  httpHeaders?: ResolvedMpvHttpHeaders;
  signal?: AbortSignal;
}) {
  const raw = await runSubtitleGenerationProcess({
    command: input.command,
    args: [
      '-v',
      'error',
      '-show_entries',
      'stream=index,codec_type,start_time,duration:stream_tags=language:format=start_time,duration,format_name',
      '-of',
      'json',
      ...(input.httpHeaders ? subtitleGenerationHttpArgs(input.httpHeaders) : []),
      input.mediaPath,
    ],
    signal: input.signal,
  });
  return parseAudioProbe(raw, input.audioStreamIndex);
}

export type SubtitleGenerationAudioProbe = Awaited<ReturnType<typeof probeSubtitleGenerationAudio>>;
