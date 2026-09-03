import test from 'node:test';
import assert from 'node:assert/strict';
import { buildIconIndex, iconMonogram, isSafeIconUrl, repoFaviconUrl } from './extension-icons';

test('buildIconIndex maps packages to their icon, skipping empty URLs', () => {
  const index = buildIconIndex([
    { pkg: 'a.b.c', iconUrl: 'https://repo.example/icon/a.b.c.png' },
    { pkg: 'd.e.f', iconUrl: '' },
  ]);
  assert.equal(index.get('a.b.c'), 'https://repo.example/icon/a.b.c.png');
  assert.equal(index.has('d.e.f'), false);
});

test('isSafeIconUrl accepts https only', () => {
  assert.equal(isSafeIconUrl('https://repo.example/icon/x.png'), true);
  assert.equal(isSafeIconUrl('http://repo.example/icon/x.png'), false);
  assert.equal(isSafeIconUrl('javascript:alert(1)'), false);
  assert.equal(isSafeIconUrl(null), false);
  assert.equal(isSafeIconUrl(undefined), false);
});

test('repoFaviconUrl points at the index host, https only', () => {
  assert.equal(
    repoFaviconUrl(' https://raw.githubusercontent.com/u/r/main/index.min.json '),
    'https://raw.githubusercontent.com/favicon.ico',
  );
  assert.equal(repoFaviconUrl('http://repo.example/index.json'), null);
  assert.equal(repoFaviconUrl('not a url'), null);
});

test('iconMonogram takes the first letter or digit, uppercased', () => {
  assert.equal(iconMonogram('AllAnime'), 'A');
  assert.equal(iconMonogram('  gogoanime'), 'G');
  assert.equal(iconMonogram('» Foo'), 'F');
  assert.equal(iconMonogram('9anime'), '9');
  assert.equal(iconMonogram('アニメ'), 'ア');
});

test('iconMonogram falls back to a placeholder for a nameless row', () => {
  assert.equal(iconMonogram(''), '?');
  assert.equal(iconMonogram('---'), '?');
});
