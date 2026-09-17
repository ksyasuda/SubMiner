import { isRemoteMediaPath } from '../jimaku/utils';
import type { MediaInput, MediaInputOptions } from '../media-input';
import type { MpvClient } from '../types/runtime';
import { resolveMpvHttpHeaders, setHeaderIfMissing } from '../core/services/mpv-http-headers';
import { extractFileUrlsFromMpvEdlSource } from './mpv-edl';

export type MediaGenerationKind = 'audio' | 'video';
export type MediaGenerationInputSource =
  | 'current-path'
  | 'stream-open-filename'
  | 'edl-stream'
  | 'youtube-cache';

export interface ResolvedMediaGenerationInput {
  path: string;
  kind: MediaGenerationKind;
  source: MediaGenerationInputSource;
  singleResolvedStream: boolean;
  inputOptions?: MediaInputOptions;
}

export interface MediaGenerationInputResolverOptions {
  getCachedMediaPath?: (
    currentVideoPath: string,
    kind: MediaGenerationKind,
  ) => Promise<string | null>;
  remoteCacheMode?: 'optional' | 'required';
  logDebug?: (message: string) => void;
}

export function resolveAudioStreamIndexForMediaGeneration(
  input: MediaInput,
  audioStreamIndex: number | null | undefined,
): number | undefined {
  if (typeof input === 'object' && 'source' in input && input.source === 'youtube-cache') {
    return undefined;
  }
  return audioStreamIndex ?? undefined;
}

function trimToNonEmptyString(value: unknown): string | null {
  if (typeof value !== 'string') {
    return null;
  }
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function extractUrlsFromMpvEdlSource(source: string): string[] {
  return extractFileUrlsFromMpvEdlSource(source)
    .map((value) => trimToNonEmptyString(value))
    .filter((value): value is string => value !== null);
}

function classifyMediaUrl(url: string): MediaGenerationKind | null {
  try {
    const mime = new URL(url).searchParams.get('mime')?.toLowerCase() ?? '';
    if (mime.startsWith('audio/')) {
      return 'audio';
    }
    if (mime.startsWith('video/')) {
      return 'video';
    }
  } catch {
    // Ignore malformed URLs and fall back to stream order.
  }

  return null;
}

function resolvePreferredUrlFromMpvEdlSource(
  source: string,
  kind: MediaGenerationKind,
): string | null {
  const urls = extractUrlsFromMpvEdlSource(source);
  if (urls.length === 0) {
    return null;
  }

  const typedMatch = urls.find((url) => classifyMediaUrl(url) === kind);
  if (typedMatch) {
    return typedMatch;
  }

  // mpv EDL sources usually list audio streams first and video streams last, so
  // when classifyMediaUrl cannot identify a typed URL we fall back to stream order.
  return kind === 'audio' ? (urls[0] ?? null) : (urls[urls.length - 1] ?? null);
}

function getHostname(value: string): string | null {
  try {
    return new URL(value).hostname.toLowerCase();
  } catch {
    return null;
  }
}

function matchesHost(hostname: string, expectedHost: string): boolean {
  return hostname === expectedHost || hostname.endsWith(`.${expectedHost}`);
}

function isGoogleVideoMediaPath(value: string): boolean {
  const host = getHostname(value);
  return Boolean(host && matchesHost(host, 'googlevideo.com'));
}

function describeMediaPathForDebugLog(value: string): string {
  try {
    const url = new URL(value);
    if (url.protocol === 'http:' || url.protocol === 'https:') {
      return `remote:${url.hostname.toLowerCase() || 'unknown'}`;
    }
    return `${url.protocol.replace(/:$/, '')}:`;
  } catch {
    // Not a URL; treat as a local file path below.
  }

  if (value.startsWith('edl://')) {
    return 'edl:';
  }

  return `local:${value}`;
}

function logMediaResolutionDebug(
  options: MediaGenerationInputResolverOptions,
  message: string,
): void {
  if (!options.logDebug) {
    return;
  }
  try {
    options.logDebug(`[media-source] ${message}`);
  } catch {
    // Debug logging should not affect media generation.
  }
}

function logResolvedMediaGenerationInput(
  options: MediaGenerationInputResolverOptions,
  currentVideoPath: string,
  result: ResolvedMediaGenerationInput,
): void {
  logMediaResolutionDebug(
    options,
    [
      `kind=${result.kind}`,
      `source=${result.source}`,
      `input=${describeMediaPathForDebugLog(result.path)}`,
      `current=${describeMediaPathForDebugLog(currentVideoPath)}`,
      `singleResolvedStream=${result.singleResolvedStream}`,
    ].join(' '),
  );
}

function logMediaGenerationInputMiss(
  options: MediaGenerationInputResolverOptions,
  kind: MediaGenerationKind,
  currentVideoPath: string,
  reason: string,
): void {
  logMediaResolutionDebug(
    options,
    [
      `kind=${kind}`,
      'source=cache-miss',
      `reason=${reason}`,
      `mode=${options.remoteCacheMode ?? 'optional'}`,
      `current=${describeMediaPathForDebugLog(currentVideoPath)}`,
    ].join(' '),
  );
}

async function resolveRemoteInputOptions(
  mpvClient: Pick<MpvClient, 'requestProperty'>,
  resolvedPath: string,
): Promise<MediaInputOptions | undefined> {
  if (!isRemoteMediaPath(resolvedPath) || !mpvClient.requestProperty) {
    return undefined;
  }

  const { headers, userAgent } = await resolveMpvHttpHeaders(mpvClient);
  if (isGoogleVideoMediaPath(resolvedPath)) {
    setHeaderIfMissing(headers, 'Referer', 'https://www.youtube.com/');
    setHeaderIfMissing(headers, 'Origin', 'https://www.youtube.com');
  }

  return {
    reconnect: true,
    ...(userAgent ? { userAgent } : {}),
    ...(Object.keys(headers).length > 0 ? { headers } : {}),
  };
}

async function toResolvedMediaGenerationInput(
  mpvClient: Pick<MpvClient, 'requestProperty'>,
  path: string,
  kind: MediaGenerationKind,
  source: MediaGenerationInputSource,
  singleResolvedStream: boolean,
): Promise<ResolvedMediaGenerationInput> {
  const inputOptions = await resolveRemoteInputOptions(mpvClient, path);
  return {
    path,
    kind,
    source,
    singleResolvedStream,
    ...(inputOptions ? { inputOptions } : {}),
  };
}

export async function resolveMediaGenerationInput(
  mpvClient: Pick<MpvClient, 'currentVideoPath' | 'requestProperty'> | null | undefined,
  kind: MediaGenerationKind = 'video',
  options: MediaGenerationInputResolverOptions = {},
): Promise<ResolvedMediaGenerationInput | null> {
  const currentVideoPath = trimToNonEmptyString(mpvClient?.currentVideoPath);
  if (!currentVideoPath) {
    logMediaResolutionDebug(options, `kind=${kind} source=none reason=no-current-video`);
    return null;
  }

  if (!isRemoteMediaPath(currentVideoPath)) {
    const result: ResolvedMediaGenerationInput = {
      path: currentVideoPath,
      kind,
      source: 'current-path',
      singleResolvedStream: false,
    };
    logResolvedMediaGenerationInput(options, currentVideoPath, result);
    return result;
  }

  let cachedPath: string | null = null;
  if (options.getCachedMediaPath) {
    try {
      cachedPath = trimToNonEmptyString(await options.getCachedMediaPath(currentVideoPath, kind));
    } catch {
      cachedPath = null;
    }
  }
  if (cachedPath) {
    const result: ResolvedMediaGenerationInput = {
      path: cachedPath,
      kind,
      source: 'youtube-cache',
      singleResolvedStream: false,
    };
    logResolvedMediaGenerationInput(options, currentVideoPath, result);
    return result;
  }

  if (options.remoteCacheMode === 'required') {
    logMediaGenerationInputMiss(options, kind, currentVideoPath, 'required-cache-unavailable');
    return null;
  }

  if (!mpvClient?.requestProperty) {
    const result: ResolvedMediaGenerationInput = {
      path: currentVideoPath,
      kind,
      source: 'current-path',
      singleResolvedStream: false,
    };
    logResolvedMediaGenerationInput(options, currentVideoPath, result);
    return result;
  }

  try {
    const streamOpenFilename = trimToNonEmptyString(
      await mpvClient.requestProperty('stream-open-filename'),
    );
    if (streamOpenFilename?.startsWith('edl://')) {
      const preferredUrl = resolvePreferredUrlFromMpvEdlSource(streamOpenFilename, kind);
      if (preferredUrl) {
        const result = await toResolvedMediaGenerationInput(
          mpvClient,
          preferredUrl,
          kind,
          'edl-stream',
          true,
        );
        logResolvedMediaGenerationInput(options, currentVideoPath, result);
        return result;
      }
      const result = await toResolvedMediaGenerationInput(
        mpvClient,
        streamOpenFilename,
        kind,
        'stream-open-filename',
        false,
      );
      logResolvedMediaGenerationInput(options, currentVideoPath, result);
      return result;
    }
    if (streamOpenFilename) {
      const result = await toResolvedMediaGenerationInput(
        mpvClient,
        streamOpenFilename,
        kind,
        'stream-open-filename',
        isRemoteMediaPath(streamOpenFilename),
      );
      logResolvedMediaGenerationInput(options, currentVideoPath, result);
      return result;
    }
  } catch {
    // Fall back to the current path when mpv does not expose a resolved stream URL.
  }

  const result = await toResolvedMediaGenerationInput(
    mpvClient,
    currentVideoPath,
    kind,
    'current-path',
    false,
  );
  logResolvedMediaGenerationInput(options, currentVideoPath, result);
  return result;
}

export async function resolveMediaGenerationInputPath(
  mpvClient: Pick<MpvClient, 'currentVideoPath' | 'requestProperty'> | null | undefined,
  kind: MediaGenerationKind = 'video',
): Promise<string | null> {
  const currentVideoPath = trimToNonEmptyString(mpvClient?.currentVideoPath);
  if (!currentVideoPath) {
    return null;
  }

  if (!isRemoteMediaPath(currentVideoPath) || !mpvClient?.requestProperty) {
    return currentVideoPath;
  }

  try {
    const streamOpenFilename = trimToNonEmptyString(
      await mpvClient.requestProperty('stream-open-filename'),
    );
    if (streamOpenFilename?.startsWith('edl://')) {
      return resolvePreferredUrlFromMpvEdlSource(streamOpenFilename, kind) ?? streamOpenFilename;
    }
    if (streamOpenFilename) {
      return streamOpenFilename;
    }
  } catch {
    // Fall back to the current path when mpv does not expose a resolved stream URL.
  }

  return currentVideoPath;
}
