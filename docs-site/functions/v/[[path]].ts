// Cloudflare Pages Function serving frozen `/v/<version>/` doc archives from R2.
// Archives are uploaded by scripts/build-versioned-docs.ts and never ship in the Pages
// deployment itself, so they do not count toward the Pages file limit. Requires an R2
// binding named DOCS_ARCHIVES on the Pages project (see docs-site/README.md).

// Minimal slice of the Workers R2 API used here; avoids a workers-types dependency.
type R2Range = { offset: number; length?: number } | { suffix: number };

type R2ObjectMeta = {
  size: number;
  httpEtag: string;
  range?: R2Range;
};

type R2ObjectBody = R2ObjectMeta & { body: ReadableStream };

type R2GetOptions = { range?: Headers; onlyIf?: Headers };

export type ArchiveBucket = {
  get(key: string, options?: R2GetOptions): Promise<R2ObjectMeta | R2ObjectBody | null>;
};

type ArchiveContext = {
  request: Request;
  env: { DOCS_ARCHIVES: ArchiveBucket };
};

export type ArchiveRoute =
  | { kind: 'redirect'; location: string; status: 301 | 302 }
  | { kind: 'lookup'; version: string; keys: string[] }
  | { kind: 'not-found' };

const CONTENT_TYPES: Record<string, string> = {
  css: 'text/css; charset=utf-8',
  gif: 'image/gif',
  html: 'text/html; charset=utf-8',
  ico: 'image/x-icon',
  jpeg: 'image/jpeg',
  jpg: 'image/jpeg',
  js: 'text/javascript; charset=utf-8',
  json: 'application/json; charset=utf-8',
  jsonc: 'application/json; charset=utf-8',
  mjs: 'text/javascript; charset=utf-8',
  mkv: 'video/x-matroska',
  mp4: 'video/mp4',
  png: 'image/png',
  svg: 'image/svg+xml',
  ttf: 'font/ttf',
  txt: 'text/plain; charset=utf-8',
  webm: 'video/webm',
  webp: 'image/webp',
  woff: 'font/woff',
  woff2: 'font/woff2',
  xml: 'application/xml; charset=utf-8',
};

function extensionOf(path: string): string | null {
  const name = path.slice(path.lastIndexOf('/') + 1);
  const dot = name.lastIndexOf('.');
  return dot > 0 ? name.slice(dot + 1).toLowerCase() : null;
}

function contentTypeFor(key: string): string {
  return CONTENT_TYPES[extensionOf(key) ?? ''] ?? 'application/octet-stream';
}

// Maps a request path onto candidate R2 keys, mirroring the Pages clean-URL rules the
// archives were built for (`cleanUrls: true`).
export function resolveArchiveRoute(pathname: string, search = ''): ArchiveRoute {
  if (pathname === '/v' || pathname === '/v/') {
    return { kind: 'redirect', location: '/versions', status: 302 };
  }

  const match = /^\/v\/(\d+\.\d+\.\d+)(\/.*)?$/.exec(pathname);
  if (!match) return { kind: 'not-found' };

  const version = match[1]!;
  const rest = match[2];
  if (!rest) {
    return { kind: 'redirect', location: `/v/${version}/${search}`, status: 301 };
  }

  let decoded: string;
  try {
    decoded = decodeURIComponent(rest);
  } catch {
    return { kind: 'not-found' };
  }
  if (decoded.split('/').some((segment) => segment === '..' || segment === '.')) {
    return { kind: 'not-found' };
  }

  const prefix = `v/${version}`;
  const path = `${prefix}${decoded}`;
  const keys = decoded.endsWith('/')
    ? [`${path}index.html`]
    : extensionOf(decoded)
      ? [path]
      : [`${path}.html`, `${path}/index.html`];

  return { kind: 'lookup', version, keys };
}

function cacheControlFor(key: string): string {
  // VitePress content-hashes everything it emits under assets/.
  if (/\/assets\//.test(key) && extensionOf(key) !== 'html') {
    return 'public, max-age=31536000, immutable';
  }
  return 'public, max-age=3600';
}

function hasBody(object: R2ObjectMeta | R2ObjectBody): object is R2ObjectBody {
  return 'body' in object && object.body !== undefined;
}

function contentRange(range: R2Range, size: number): { start: number; end: number } {
  if ('suffix' in range) {
    const length = Math.min(range.suffix, size);
    return { start: size - length, end: size - 1 };
  }
  const length = range.length ?? size - range.offset;
  return { start: range.offset, end: range.offset + length - 1 };
}

async function respondWithObject(options: {
  request: Request;
  bucket: ArchiveBucket;
  key: string;
  status: number;
}): Promise<Response | null> {
  const { request, bucket, key } = options;
  // Ranges and conditional requests only make sense for the page that was asked for,
  // not the 404 fallback.
  const isRequestedObject = options.status === 200;
  const wantsRange = isRequestedObject && request.headers.has('range');

  let object: R2ObjectMeta | R2ObjectBody | null;
  try {
    object = await bucket.get(key, {
      range: wantsRange ? request.headers : undefined,
      onlyIf: isRequestedObject ? request.headers : undefined,
    });
  } catch {
    // R2 rejects unsatisfiable ranges.
    return new Response(null, { status: 416 });
  }
  if (!object) return null;

  const headers = new Headers({
    'Content-Type': contentTypeFor(key),
    'Cache-Control': cacheControlFor(key),
    ETag: object.httpEtag,
    'Accept-Ranges': 'bytes',
    'X-Robots-Tag': 'noindex, follow',
  });

  // R2 returns the object without a body when an If-None-Match/If-Modified-Since
  // precondition matched.
  if (!hasBody(object)) {
    return new Response(null, { status: 304, headers });
  }

  let status = options.status;
  if (wantsRange && object.range) {
    const { start, end } = contentRange(object.range, object.size);
    status = 206;
    headers.set('Content-Range', `bytes ${start}-${end}/${object.size}`);
    headers.set('Content-Length', String(end - start + 1));
  } else {
    headers.set('Content-Length', String(object.size));
  }

  return new Response(request.method === 'HEAD' ? null : object.body, { status, headers });
}

export async function onRequest({ request, env }: ArchiveContext): Promise<Response> {
  if (request.method !== 'GET' && request.method !== 'HEAD') {
    return new Response('Method Not Allowed', { status: 405, headers: { Allow: 'GET, HEAD' } });
  }

  const url = new URL(request.url);
  const route = resolveArchiveRoute(url.pathname, url.search);

  if (route.kind === 'redirect') {
    return Response.redirect(new URL(route.location, url).toString(), route.status);
  }

  if (route.kind === 'lookup') {
    for (const key of route.keys) {
      const response = await respondWithObject({
        request,
        bucket: env.DOCS_ARCHIVES,
        key,
        status: 200,
      });
      if (response) return response;
    }

    const notFoundPage = await respondWithObject({
      request,
      bucket: env.DOCS_ARCHIVES,
      key: `v/${route.version}/404.html`,
      status: 404,
    });
    if (notFoundPage) return notFoundPage;
  }

  return new Response('Not Found', {
    status: 404,
    headers: { 'Content-Type': 'text/plain; charset=utf-8', 'X-Robots-Tag': 'noindex, follow' },
  });
}
