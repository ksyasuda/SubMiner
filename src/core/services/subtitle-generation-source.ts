import { toFfmpegInputHttpArgs, type ResolvedMpvHttpHeaders } from './mpv-http-headers';

export interface SubtitleGenerationRemoteSource {
  httpHeaders: ResolvedMpvHttpHeaders;
  cacheDirectory: string;
  sessionDirectory?: string;
  alternatives?: readonly SubtitleGenerationAlternative[];
  selectedAudio?: {
    url: string;
    audioStreamIndex?: number;
    delaySeconds: number;
    httpHeaders: ResolvedMpvHttpHeaders;
  };
}

export interface SubtitleGenerationAlternative {
  kind: 'audio' | 'video';
  url: string;
  label: string;
  httpHeaders: ResolvedMpvHttpHeaders;
}

/** Apply only to network inputs, including embedded subtitle reference extraction. */
export function subtitleGenerationHttpArgs(headers: ResolvedMpvHttpHeaders): string[] {
  return [
    '-protocol_whitelist',
    'http,https,tcp,tls,crypto',
    '-rw_timeout',
    '15000000',
    ...toFfmpegInputHttpArgs(headers),
  ];
}
