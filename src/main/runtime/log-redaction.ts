import * as net from 'net';

type LogRedactionExample = {
  input: string;
  expected: string;
};

export type LogRedactionRule = {
  name: string;
  pattern: string;
  replacement: string;
  risk: 'pii' | 'secret' | 'pii-or-secret';
  examples: readonly LogRedactionExample[];
  redact: (text: string) => string;
};

const REDACTED_USER = '<user>';
const REDACTED_VALUE = '<redacted>';
const REDACTED_CREDENTIALS = '<credentials>';
const REDACTED_EMAIL = '<email>';
const REDACTED_IP = '<ip>';

const SENSITIVE_HEADER_NAMES = [
  'api-key',
  'authorization',
  'cookie',
  'proxy-authorization',
  'set-cookie',
  'x-api-key',
  'x-emby-token',
  'x-goog-visitor-id',
  'x-mediabrowser-token',
].join('|');

const SENSITIVE_VALUE_KEY_PATTERNS = [
  'access[_-]?token',
  'api[_-]?key',
  'apikey',
  'authorization',
  'client[_-]?secret',
  'cookie',
  'cookies',
  'id[_-]?token',
  'password',
  'passwd',
  'pwd',
  'refresh[_-]?token',
  'secret',
  'session',
  'sid',
  'sig',
  'signature',
  'token',
].join('|');

const SENSITIVE_VALUE_KEY_RE = new RegExp(`^(?:${SENSITIVE_VALUE_KEY_PATTERNS})$`, 'i');
const SENSITIVE_HEADER_RE = new RegExp(
  `(^|\\r?\\n)(\\s*(?:${SENSITIVE_HEADER_NAMES})\\s*:\\s*)[^\\r\\n]*`,
  'gi',
);
const SENSITIVE_INLINE_HEADER_RE = new RegExp(
  `\\b((?:${SENSITIVE_HEADER_NAMES})\\s*:\\s*)[^\\r\\n]*`,
  'gi',
);
const SENSITIVE_QUOTED_VALUE_RE = new RegExp(
  `(["'])(${SENSITIVE_VALUE_KEY_PATTERNS})\\1(\\s*[:=]\\s*)(["'])([^"'\\r\\n]*)\\4`,
  'gi',
);
const SENSITIVE_UNQUOTED_VALUE_RE = new RegExp(
  `\\b(${SENSITIVE_VALUE_KEY_PATTERNS})(\\s*[:=]\\s*)(?:"[^"\\r\\n]*"|'[^'\\r\\n]*'|[^\\s,}\\]\\)&<>"']+)`,
  'gi',
);
const YTDLP_COOKIE_EQUALS_RE =
  /(--(?:cookies|cookies-from-browser)=)(?:"[^"\r\n]*"|'[^'\r\n]*'|[^\s\r\n]+)/gi;
const YTDLP_COOKIE_SPACE_RE =
  /(--(?:cookies|cookies-from-browser)\s+)(?:"[^"\r\n]*"|'[^'\r\n]*'|[^\r\n]*?)(?=\s+--[A-Za-z0-9][A-Za-z0-9-]*|\r?\n|$)/gi;
const URL_TOKEN_RE = /\b[a-z][a-z0-9+.-]*:\/\/[^\s"'<>]+/gi;
const URL_CREDENTIALS_RE = /\b([a-z][a-z0-9+.-]*:\/\/)([^/@\s]+@)/gi;
const URL_QUERY_PAIR_RE = /([?&;]|&amp;)([^=&#;]+)=([^&#;]*)/gi;
const TRAILING_URL_PUNCTUATION_RE = /[)\].,;:!?]+$/;
const EMAIL_RE = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
const BRACKETED_IP_RE = /\[([0-9A-F:.%]+)\]/gi;
const IPV4_RE = /(^|[^A-Za-z0-9_.-])((?:\d{1,3}\.){3}\d{1,3})(?![A-Za-z0-9_.-])/g;
const IPV6_RE = /(^|[^A-Za-z0-9_.-])([0-9A-F]{0,4}:[0-9A-F:.%]{2,})(?![A-Za-z0-9_.-])/gi;

function safeDecodeFormComponent(value: string): string {
  try {
    return decodeURIComponent(value.replace(/\+/g, ' '));
  } catch {
    return value;
  }
}

function isSensitiveValueKey(key: string): boolean {
  return SENSITIVE_VALUE_KEY_RE.test(safeDecodeFormComponent(key));
}

function splitTrailingUrlPunctuation(token: string): { core: string; trailing: string } {
  const match = token.match(TRAILING_URL_PUNCTUATION_RE);
  if (!match) return { core: token, trailing: '' };
  return {
    core: token.slice(0, -match[0].length),
    trailing: match[0],
  };
}

function redactSensitiveUrlQueryPairs(rawUrl: string): string {
  const queryStart = rawUrl.indexOf('?');
  if (queryStart === -1) return rawUrl;

  const hashStart = rawUrl.indexOf('#', queryStart);
  const queryEnd = hashStart === -1 ? rawUrl.length : hashStart;
  const query = rawUrl.slice(queryStart, queryEnd);
  const redactedQuery = query.replace(URL_QUERY_PAIR_RE, (match, prefix: string, rawKey: string) =>
    isSensitiveValueKey(rawKey) ? `${prefix}${rawKey}=${REDACTED_VALUE}` : match,
  );

  return `${rawUrl.slice(0, queryStart)}${redactedQuery}${rawUrl.slice(queryEnd)}`;
}

function redactUrlToken(token: string): string {
  const { core, trailing } = splitTrailingUrlPunctuation(token);
  try {
    new URL(core);
  } catch {
    return token;
  }

  const withoutCredentials = core.replace(URL_CREDENTIALS_RE, `$1${REDACTED_CREDENTIALS}@`);
  return `${redactSensitiveUrlQueryPairs(withoutCredentials)}${trailing}`;
}

function redactUrlSensitiveComponents(text: string): string {
  return text.replace(URL_TOKEN_RE, (token: string) => redactUrlToken(token));
}

function isGoogleVideoHost(hostname: string): boolean {
  const normalized = hostname.toLowerCase();
  return normalized === 'googlevideo.com' || normalized.endsWith('.googlevideo.com');
}

function isSignedYouTubeMediaUrl(url: URL): boolean {
  return isGoogleVideoHost(url.hostname) && url.pathname === '/videoplayback' && url.search !== '';
}

function redactSignedYouTubeMediaUrlToken(token: string): string {
  const { core, trailing } = splitTrailingUrlPunctuation(token);
  let parsed: URL;
  try {
    parsed = new URL(core);
  } catch {
    return token;
  }

  if (!isSignedYouTubeMediaUrl(parsed)) return token;

  const queryStart = core.indexOf('?');
  const hashStart = core.indexOf('#', queryStart);
  if (queryStart === -1) return token;

  const hash = hashStart === -1 ? '' : core.slice(hashStart);
  return `${core.slice(0, queryStart)}?${REDACTED_VALUE}${hash}${trailing}`;
}

function redactSignedYouTubeMediaUrlQueries(text: string): string {
  return text.replace(URL_TOKEN_RE, (token: string) => redactSignedYouTubeMediaUrlToken(token));
}

function redactSensitiveHeaders(text: string): string {
  return text
    .replace(
      SENSITIVE_HEADER_RE,
      (_match, lineStart: string, headerPrefix: string) =>
        `${lineStart}${headerPrefix}${REDACTED_VALUE}`,
    )
    .replace(
      SENSITIVE_INLINE_HEADER_RE,
      (_match, headerPrefix: string) => `${headerPrefix}${REDACTED_VALUE}`,
    );
}

function redactYtDlpCookieArgs(text: string): string {
  return text
    .replace(YTDLP_COOKIE_EQUALS_RE, `$1${REDACTED_VALUE}`)
    .replace(YTDLP_COOKIE_SPACE_RE, `$1${REDACTED_VALUE}`);
}

function findJsonEnd(text: string, start: number): number {
  const stack: string[] = [];
  let inString = false;
  let quote = '';
  let escaped = false;

  for (let index = start; index < text.length; index += 1) {
    const char = text[index];

    if (inString) {
      if (escaped) {
        escaped = false;
      } else if (char === '\\') {
        escaped = true;
      } else if (char === quote) {
        inString = false;
      }
      continue;
    }

    if (char === '"' || char === "'") {
      inString = true;
      quote = char;
      continue;
    }

    if (char === '{') {
      stack.push('}');
    } else if (char === '[') {
      stack.push(']');
    } else if (char === '}' || char === ']') {
      if (stack.pop() !== char) return -1;
      if (stack.length === 0) return index;
    }
  }

  return -1;
}

function redactJsonValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map((entry) => redactJsonValue(entry));
  }

  if (value && typeof value === 'object') {
    const redacted: Record<string, unknown> = {};
    for (const [key, entry] of Object.entries(value)) {
      redacted[key] = isSensitiveValueKey(key) ? REDACTED_VALUE : redactJsonValue(entry);
    }
    return redacted;
  }

  return value;
}

function redactJsonPayloads(text: string): string {
  let output = '';
  let cursor = 0;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (char !== '{' && char !== '[') continue;

    const end = findJsonEnd(text, index);
    if (end === -1) continue;

    const candidate = text.slice(index, end + 1);
    try {
      const parsed = JSON.parse(candidate);
      output += text.slice(cursor, index);
      output += JSON.stringify(redactJsonValue(parsed));
      cursor = end + 1;
      index = end;
    } catch {
      // Non-JSON brace pairs remain available to the fallback key-value rules.
    }
  }

  return `${output}${text.slice(cursor)}`;
}

function redactSensitiveKeyValues(text: string): string {
  return text
    .replace(
      SENSITIVE_QUOTED_VALUE_RE,
      (_match, keyQuote: string, key: string, separator: string, valueQuote: string) =>
        `${keyQuote}${key}${keyQuote}${separator}${valueQuote}${REDACTED_VALUE}${valueQuote}`,
    )
    .replace(SENSITIVE_UNQUOTED_VALUE_RE, (match, key: string, separator: string) => {
      const value = match.slice(key.length + separator.length);
      const quote = value[0];
      if (quote === '"' || quote === "'") {
        return `${key}${separator}${quote}${REDACTED_VALUE}${quote}`;
      }
      return `${key}${separator}${REDACTED_VALUE}`;
    });
}

function redactHomePathUsernames(text: string): string {
  return text
    .replace(/(\/(?:home|Users)\/)([^/\r\n]+)(?=\/|$)/g, `$1${REDACTED_USER}`)
    .replace(/([A-Za-z]:[\\/]+Users[\\/]+)([^\\/:\r\n]+)(?=[\\/]|$)/g, `$1${REDACTED_USER}`);
}

function normalizeIpCandidate(candidate: string): string {
  return candidate.split('%', 1)[0] ?? candidate;
}

function isIpAddress(candidate: string): boolean {
  return net.isIP(normalizeIpCandidate(candidate)) !== 0;
}

function redactIpAddresses(text: string): string {
  return text
    .replace(BRACKETED_IP_RE, (match, candidate: string) =>
      isIpAddress(candidate) ? `[${REDACTED_IP}]` : match,
    )
    .replace(IPV4_RE, (match, prefix: string, candidate: string) =>
      net.isIP(candidate) === 4 ? `${prefix}${REDACTED_IP}` : match,
    )
    .replace(IPV6_RE, (match, prefix: string, candidate: string) =>
      net.isIP(normalizeIpCandidate(candidate)) === 6 ? `${prefix}${REDACTED_IP}` : match,
    );
}

export const LOG_EXPORT_REDACTION_RULES: readonly LogRedactionRule[] = [
  {
    name: 'signed-youtube-media-url-query',
    pattern: '*.googlevideo.com/videoplayback query strings',
    replacement: REDACTED_VALUE,
    risk: 'pii-or-secret',
    examples: [
      {
        input: 'https://rr1---sn.example.googlevideo.com/videoplayback?expire=1&sig=secret',
        expected: `https://rr1---sn.example.googlevideo.com/videoplayback?${REDACTED_VALUE}`,
      },
    ],
    redact: redactSignedYouTubeMediaUrlQueries,
  },
  {
    name: 'url-sensitive-components',
    pattern: 'URL tokens with credentials or sensitive query params',
    replacement: `${REDACTED_CREDENTIALS}, ${REDACTED_VALUE}`,
    risk: 'pii-or-secret',
    examples: [
      {
        input: 'GET https://alice:secret@example.test/watch?access_token=tok123&state=ok',
        expected: `GET https://${REDACTED_CREDENTIALS}@example.test/watch?access_token=${REDACTED_VALUE}&state=ok`,
      },
    ],
    redact: redactUrlSensitiveComponents,
  },
  {
    name: 'sensitive-headers',
    pattern: `headers: ${SENSITIVE_HEADER_NAMES}`,
    replacement: REDACTED_VALUE,
    risk: 'secret',
    examples: [
      {
        input: 'Authorization: Bearer token',
        expected: `Authorization: ${REDACTED_VALUE}`,
      },
    ],
    redact: redactSensitiveHeaders,
  },
  {
    name: 'yt-dlp-cookie-args',
    pattern: '--cookies and --cookies-from-browser arguments',
    replacement: REDACTED_VALUE,
    risk: 'secret',
    examples: [
      {
        input: 'yt-dlp --cookies /Users/kyle/cookies.txt --verbose',
        expected: `yt-dlp --cookies ${REDACTED_VALUE} --verbose`,
      },
    ],
    redact: redactYtDlpCookieArgs,
  },
  {
    name: 'json-secret-values',
    pattern: `JSON object keys: ${SENSITIVE_VALUE_KEY_PATTERNS}`,
    replacement: REDACTED_VALUE,
    risk: 'secret',
    examples: [
      {
        input: 'json {"token":123,"safe":"ok"}',
        expected: `json {"token":"${REDACTED_VALUE}","safe":"ok"}`,
      },
    ],
    redact: redactJsonPayloads,
  },
  {
    name: 'generic-sensitive-key-values',
    pattern: `key-value pairs: ${SENSITIVE_VALUE_KEY_PATTERNS}`,
    replacement: REDACTED_VALUE,
    risk: 'secret',
    examples: [
      {
        input: 'refreshToken=abc123 state=ok',
        expected: `refreshToken=${REDACTED_VALUE} state=ok`,
      },
    ],
    redact: redactSensitiveKeyValues,
  },
  {
    name: 'home-path-usernames',
    pattern: 'Linux, macOS, and Windows home paths',
    replacement: REDACTED_USER,
    risk: 'pii',
    examples: [
      {
        input: '/Users/kyle/Library/Application Support/SubMiner',
        expected: `/Users/${REDACTED_USER}/Library/Application Support/SubMiner`,
      },
    ],
    redact: redactHomePathUsernames,
  },
  {
    name: 'email-addresses',
    pattern: 'email-like addresses',
    replacement: REDACTED_EMAIL,
    risk: 'pii',
    examples: [
      {
        input: 'support kyle@example.test',
        expected: `support ${REDACTED_EMAIL}`,
      },
    ],
    redact: (text) => text.replace(EMAIL_RE, REDACTED_EMAIL),
  },
  {
    name: 'ip-addresses',
    pattern: 'IPv4, bracketed IPv6, and bare IPv6 candidates validated with net.isIP',
    replacement: REDACTED_IP,
    risk: 'pii',
    examples: [
      {
        input: 'remote addr [2001:db8::1234]:443 and 203.0.113.42',
        expected: `remote addr [${REDACTED_IP}]:443 and ${REDACTED_IP}`,
      },
    ],
    redact: redactIpAddresses,
  },
];

export function redactLogExportText(text: string): string {
  return LOG_EXPORT_REDACTION_RULES.reduce((redacted, rule) => rule.redact(redacted), text);
}
