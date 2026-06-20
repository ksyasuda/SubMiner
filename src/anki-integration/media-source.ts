import { isRemoteMediaPath } from '../jimaku/utils';
import type { MediaInputOptions } from '../media-input';
import type { MpvClient } from '../types/runtime';

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
}

const BLOCKED_HTTP_HEADER_NAMES = new Set(['authorization', 'cookie', 'proxy-authorization']);
const HTTP_HEADER_FIELD_PROPERTY_NAMES = [
  'http-header-fields',
  'options/http-header-fields',
  'file-local-options/http-header-fields',
] as const;
const USER_AGENT_PROPERTY_NAMES = [
  'file-local-options/user-agent',
  'options/user-agent',
  'user-agent',
] as const;
const REFERRER_PROPERTY_NAMES = [
  'file-local-options/referrer',
  'options/referrer',
  'referrer',
] as const;

function trimToNonEmptyString(value: unknown): string | null {
  if (typeof value !== 'string') {
    return null;
  }
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function normalizeHeaderName(value: string): string | null {
  const trimmed = value.trim();
  if (!/^[A-Za-z0-9!#$%&'*+.^_`|~-]+$/.test(trimmed)) {
    return null;
  }
  if (BLOCKED_HTTP_HEADER_NAMES.has(trimmed.toLowerCase())) {
    return null;
  }
  return trimmed;
}

function extractUrlsFromMpvEdlSource(source: string): string[] {
  const matches = source.matchAll(/%\d+%(https?:\/\/.*?)(?=;!new_stream|;!global_tags|$)/gms);
  return [...matches]
    .map((match) => trimToNonEmptyString(match[1]))
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

function setHeaderIfMissing(headers: Record<string, string>, name: string, value: string): void {
  const lowerName = name.toLowerCase();
  if (!Object.keys(headers).some((existing) => existing.toLowerCase() === lowerName)) {
    headers[name] = value;
  }
}

function parseMpvHeaderField(value: string): [string, string] | null {
  const separatorIndex = value.indexOf(':');
  if (separatorIndex <= 0) {
    return null;
  }

  const name = normalizeHeaderName(value.slice(0, separatorIndex));
  const headerValue = trimToNonEmptyString(value.slice(separatorIndex + 1));
  if (!name || !headerValue) {
    return null;
  }
  return [name, headerValue.replace(/[\r\n]+/g, ' ')];
}

function toHeaderFields(value: unknown): string[] {
  if (Array.isArray(value)) {
    return value.filter((entry): entry is string => typeof entry === 'string');
  }
  if (typeof value === 'string') {
    return value
      .split(/\r?\n/)
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0);
  }
  return [];
}

async function requestOptionalMpvProperty(
  mpvClient: Pick<MpvClient, 'requestProperty'>,
  name: string,
): Promise<unknown> {
  if (!mpvClient.requestProperty) {
    return null;
  }

  try {
    return await mpvClient.requestProperty(name);
  } catch {
    return null;
  }
}

async function requestFirstNonEmptyStringProperty(
  mpvClient: Pick<MpvClient, 'requestProperty'>,
  names: readonly string[],
): Promise<string | null> {
  for (const name of names) {
    const value = trimToNonEmptyString(await requestOptionalMpvProperty(mpvClient, name));
    if (value) {
      return value;
    }
  }
  return null;
}

async function resolveRemoteInputOptions(
  mpvClient: Pick<MpvClient, 'requestProperty'>,
  resolvedPath: string,
): Promise<MediaInputOptions | undefined> {
  if (!isRemoteMediaPath(resolvedPath) || !mpvClient.requestProperty) {
    return undefined;
  }

  const headers: Record<string, string> = {};
  for (const propertyName of HTTP_HEADER_FIELD_PROPERTY_NAMES) {
    const mpvHeaderFields = toHeaderFields(
      await requestOptionalMpvProperty(mpvClient, propertyName),
    );
    for (const field of mpvHeaderFields) {
      const parsed = parseMpvHeaderField(field);
      if (parsed) {
        headers[parsed[0]] = parsed[1];
      }
    }
  }

  const userAgent = await requestFirstNonEmptyStringProperty(mpvClient, USER_AGENT_PROPERTY_NAMES);
  const referrer = await requestFirstNonEmptyStringProperty(mpvClient, REFERRER_PROPERTY_NAMES);
  if (referrer) {
    setHeaderIfMissing(headers, 'Referer', referrer);
  }
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
    return null;
  }

  if (!isRemoteMediaPath(currentVideoPath)) {
    return {
      path: currentVideoPath,
      kind,
      source: 'current-path',
      singleResolvedStream: false,
    };
  }

  const cachedPath = options.getCachedMediaPath
    ? trimToNonEmptyString(await options.getCachedMediaPath(currentVideoPath, kind))
    : null;
  if (cachedPath) {
    return {
      path: cachedPath,
      kind,
      source: 'youtube-cache',
      singleResolvedStream: false,
    };
  }

  if (!mpvClient?.requestProperty) {
    return {
      path: currentVideoPath,
      kind,
      source: 'current-path',
      singleResolvedStream: false,
    };
  }

  try {
    const streamOpenFilename = trimToNonEmptyString(
      await mpvClient.requestProperty('stream-open-filename'),
    );
    if (streamOpenFilename?.startsWith('edl://')) {
      const preferredUrl = resolvePreferredUrlFromMpvEdlSource(streamOpenFilename, kind);
      if (preferredUrl) {
        return toResolvedMediaGenerationInput(mpvClient, preferredUrl, kind, 'edl-stream', true);
      }
      return toResolvedMediaGenerationInput(
        mpvClient,
        streamOpenFilename,
        kind,
        'stream-open-filename',
        false,
      );
    }
    if (streamOpenFilename) {
      return toResolvedMediaGenerationInput(
        mpvClient,
        streamOpenFilename,
        kind,
        'stream-open-filename',
        isRemoteMediaPath(streamOpenFilename),
      );
    }
  } catch {
    // Fall back to the current path when mpv does not expose a resolved stream URL.
  }

  return toResolvedMediaGenerationInput(mpvClient, currentVideoPath, kind, 'current-path', false);
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
