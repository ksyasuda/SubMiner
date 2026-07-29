import assert from 'node:assert/strict';
import test from 'node:test';
import { assetUrl, resolveAssetUrl } from './asset-url';

// vite.config.ts sets `base: './'`, so this is what the built bundle sees.
const BUILT_BASE = './';

test('built asset URLs are never root-absolute', () => {
  assert.equal(resolveAssetUrl('favicon.png', BUILT_BASE).startsWith('/'), false);
});

test('built asset URL resolves next to a file:// index.html', () => {
  const resolved = new URL(
    resolveAssetUrl('favicon.png', BUILT_BASE),
    'file:///opt/SubMiner/stats/dist/index.html',
  );
  assert.equal(resolved.href, 'file:///opt/SubMiner/stats/dist/favicon.png');
});

test('built asset URL resolves against the server root when served over http', () => {
  const resolved = new URL(resolveAssetUrl('favicon.png', BUILT_BASE), 'http://127.0.0.1:8770/');
  assert.equal(resolved.href, 'http://127.0.0.1:8770/favicon.png');
});

test('dev server base stays root-absolute', () => {
  assert.equal(resolveAssetUrl('favicon.png', '/'), '/favicon.png');
});

test('a base without a trailing slash still joins cleanly', () => {
  assert.equal(resolveAssetUrl('favicon.png', '/stats'), '/stats/favicon.png');
});

test('a leading slash in the requested path is tolerated', () => {
  assert.equal(resolveAssetUrl('/favicon.png', BUILT_BASE), './favicon.png');
});

test('assetUrl falls back to a relative base outside a Vite bundle', () => {
  assert.equal(assetUrl('favicon.png').startsWith('/'), false);
});
