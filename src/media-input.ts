export interface MediaInputOptions {
  reconnect?: boolean;
  userAgent?: string;
  headers?: Record<string, string>;
}

export type MediaInput =
  | string
  | {
      path: string;
      source?: string;
      inputOptions?: MediaInputOptions;
      singleResolvedStream?: boolean;
    };

export type NormalizedMediaInput = {
  path: string;
  inputArgs: string[];
  singleResolvedStream: boolean;
};

const BLOCKED_FFMPEG_HEADER_NAMES = new Set(['authorization', 'cookie', 'proxy-authorization']);

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
  if (BLOCKED_FFMPEG_HEADER_NAMES.has(trimmed.toLowerCase())) {
    return null;
  }
  return trimmed;
}

function serializeFfmpegHeaders(headers: Record<string, string> | undefined): string | null {
  if (!headers) {
    return null;
  }

  const lines: string[] = [];
  for (const [rawName, rawValue] of Object.entries(headers)) {
    const name = normalizeHeaderName(rawName);
    const value = trimToNonEmptyString(rawValue);
    if (!name || !value) {
      continue;
    }
    lines.push(`${name}: ${value.replace(/[\r\n]+/g, ' ')}`);
  }

  return lines.length > 0 ? `${lines.join('\r\n')}\r\n` : null;
}

export function normalizeMediaInput(input: MediaInput): NormalizedMediaInput {
  if (typeof input === 'string') {
    return { path: input, inputArgs: [], singleResolvedStream: false };
  }

  const inputArgs: string[] = [];
  if (input.inputOptions?.reconnect) {
    inputArgs.push('-reconnect', '1', '-reconnect_streamed', '1', '-reconnect_delay_max', '5');
  }

  const userAgent = trimToNonEmptyString(input.inputOptions?.userAgent);
  if (userAgent) {
    inputArgs.push('-user_agent', userAgent);
  }

  const headers = serializeFfmpegHeaders(input.inputOptions?.headers);
  if (headers) {
    inputArgs.push('-headers', headers);
  }

  return {
    path: input.path,
    inputArgs,
    singleResolvedStream: input.singleResolvedStream === true,
  };
}
