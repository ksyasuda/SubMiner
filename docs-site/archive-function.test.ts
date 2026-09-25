import { expect, test } from 'bun:test';
import { onRequest, resolveArchiveRoute, type ArchiveBucket } from './functions/v/[[path]]';

// In-memory stand-in for the R2 binding, including the range and precondition behavior
// the function relies on.
function fakeBucket(objects: Record<string, string>): ArchiveBucket {
  return {
    async get(key, options) {
      const content = objects[key];
      if (content === undefined) return null;
      const bytes = new TextEncoder().encode(content);
      const httpEtag = `"${key}"`;

      if (options?.onlyIf?.get('if-none-match') === httpEtag) {
        return { size: bytes.length, httpEtag };
      }

      const rangeHeader = options?.range?.get('range');
      const match = rangeHeader ? /^bytes=(\d+)-(\d*)$/.exec(rangeHeader) : null;
      if (match) {
        const offset = Number(match[1]);
        const end = match[2] ? Number(match[2]) : bytes.length - 1;
        const slice = bytes.slice(offset, end + 1);
        return {
          size: bytes.length,
          httpEtag,
          range: { offset, length: slice.length },
          body: new Blob([slice]).stream(),
        };
      }

      return { size: bytes.length, httpEtag, body: new Blob([bytes]).stream() };
    },
  };
}

const bucket = fakeBucket({
  'v/0.19.6/index.html': '<h1>home</h1>',
  'v/0.19.6/usage.html': '<h1>usage</h1>',
  'v/0.19.6/404.html': '<h1>missing</h1>',
  'v/0.19.6/assets/app.abc123.js': 'console.log(1)',
});

function request(path: string, init?: RequestInit) {
  return onRequest({
    request: new Request(`https://docs.subminer.moe${path}`, init),
    env: { DOCS_ARCHIVES: bucket },
  });
}

test('archive routes follow the clean-URL layout of the built archives', () => {
  expect(resolveArchiveRoute('/v/0.19.6/usage')).toEqual({
    kind: 'lookup',
    version: '0.19.6',
    keys: ['v/0.19.6/usage.html', 'v/0.19.6/usage/index.html'],
  });
  expect(resolveArchiveRoute('/v/0.19.6/')).toEqual({
    kind: 'lookup',
    version: '0.19.6',
    keys: ['v/0.19.6/index.html'],
  });
  expect(resolveArchiveRoute('/v/0.19.6', '?q=1')).toEqual({
    kind: 'redirect',
    location: '/v/0.19.6/?q=1',
    status: 301,
  });
  expect(resolveArchiveRoute('/v/')).toEqual({
    kind: 'redirect',
    location: '/versions',
    status: 302,
  });
  expect(resolveArchiveRoute('/v/0.19.6/%2e%2e/secret')).toEqual({ kind: 'not-found' });
  expect(resolveArchiveRoute('/v/latest/')).toEqual({ kind: 'not-found' });
});

test('serves archive pages and hashed assets with their cache policy', async () => {
  const page = await request('/v/0.19.6/usage');
  expect(page.status).toBe(200);
  expect(page.headers.get('content-type')).toBe('text/html; charset=utf-8');
  expect(page.headers.get('cache-control')).toBe('public, max-age=3600');
  expect(page.headers.get('x-robots-tag')).toBe('noindex, follow');
  expect(await page.text()).toBe('<h1>usage</h1>');

  const asset = await request('/v/0.19.6/assets/app.abc123.js');
  expect(asset.headers.get('content-type')).toBe('text/javascript; charset=utf-8');
  expect(asset.headers.get('cache-control')).toContain('immutable');
});

test('missing archive pages fall back to the archive 404 page', async () => {
  const response = await request('/v/0.19.6/nope');
  expect(response.status).toBe(404);
  expect(await response.text()).toBe('<h1>missing</h1>');

  const unknownVersion = await request('/v/0.1.0/usage');
  expect(unknownVersion.status).toBe(404);
});

test('supports byte ranges, conditional requests, and HEAD', async () => {
  const partial = await request('/v/0.19.6/usage', { headers: { Range: 'bytes=4-8' } });
  expect(partial.status).toBe(206);
  expect(partial.headers.get('content-range')).toBe('bytes 4-8/14');
  expect(await partial.text()).toBe('usage');

  const notModified = await request('/v/0.19.6/usage', {
    headers: { 'If-None-Match': '"v/0.19.6/usage.html"' },
  });
  expect(notModified.status).toBe(304);

  const head = await request('/v/0.19.6/usage', { method: 'HEAD' });
  expect(head.status).toBe(200);
  expect(head.headers.get('content-length')).toBe('14');
  expect(await head.text()).toBe('');

  const post = await request('/v/0.19.6/usage', { method: 'POST' });
  expect(post.status).toBe(405);
});
