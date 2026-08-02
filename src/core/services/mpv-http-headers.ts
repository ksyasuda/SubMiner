/**
 * Read back the HTTP request options mpv is using for the current stream.
 *
 * Anything SubMiner fetches out-of-band — ffmpeg extractions, subtitle
 * downloads — has to speak to the origin the same way mpv did, or an
 * authenticated or referer-gated host answers 403 to us while playback works.
 */

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

export interface MpvHttpPropertySource {
  requestProperty?: (name: string) => Promise<unknown>;
}

export interface ResolvedMpvHttpHeaders {
  headers: Record<string, string>;
  userAgent: string | null;
}

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

export function setHeaderIfMissing(
  headers: Record<string, string>,
  name: string,
  value: string,
): void {
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
  client: MpvHttpPropertySource,
  name: string,
): Promise<unknown> {
  if (!client.requestProperty) {
    return null;
  }

  try {
    return await client.requestProperty(name);
  } catch {
    return null;
  }
}

async function requestFirstNonEmptyStringProperty(
  client: MpvHttpPropertySource,
  names: readonly string[],
): Promise<string | null> {
  for (const name of names) {
    const value = trimToNonEmptyString(await requestOptionalMpvProperty(client, name));
    if (value) {
      return value;
    }
  }
  return null;
}

export async function resolveMpvHttpHeaders(
  client: MpvHttpPropertySource,
): Promise<ResolvedMpvHttpHeaders> {
  const headers: Record<string, string> = {};
  if (!client.requestProperty) {
    return { headers, userAgent: null };
  }

  for (const propertyName of HTTP_HEADER_FIELD_PROPERTY_NAMES) {
    const mpvHeaderFields = toHeaderFields(await requestOptionalMpvProperty(client, propertyName));
    for (const field of mpvHeaderFields) {
      const parsed = parseMpvHeaderField(field);
      if (parsed) {
        headers[parsed[0]] = parsed[1];
      }
    }
  }

  const userAgent = await requestFirstNonEmptyStringProperty(client, USER_AGENT_PROPERTY_NAMES);
  const referrer = await requestFirstNonEmptyStringProperty(client, REFERRER_PROPERTY_NAMES);
  if (referrer) {
    setHeaderIfMissing(headers, 'Referer', referrer);
  }

  return { headers, userAgent };
}

/** Flatten to the header map a plain `http.get` wants. */
export function toRequestHeaders(resolved: ResolvedMpvHttpHeaders | null): Record<string, string> {
  if (!resolved) return {};
  const headers = { ...resolved.headers };
  if (resolved.userAgent) {
    setHeaderIfMissing(headers, 'User-Agent', resolved.userAgent);
  }
  return headers;
}

/**
 * ffmpeg input options carrying the same request context. These configure the
 * HTTP demuxer, so they must precede `-i` on the command line.
 */
export function toFfmpegInputHttpArgs(resolved: ResolvedMpvHttpHeaders | null): string[] {
  if (!resolved) return [];

  const args: string[] = [];
  if (resolved.userAgent) {
    args.push('-user_agent', resolved.userAgent);
  }
  const fields = Object.entries(resolved.headers);
  if (fields.length > 0) {
    args.push('-headers', fields.map(([name, value]) => `${name}: ${value}\r\n`).join(''));
  }
  return args;
}
