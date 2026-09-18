import assert from 'node:assert/strict';
import test from 'node:test';
import { sanitizeMediaTitle, toMediaIdentityPath } from './media-identity';

test('media identity separates authenticated transport URLs from titles and stats keys', () => {
  const url =
    'https://user:password@jellyfin.example/Videos/item-1/stream?api_key=test-secret#token';
  assert.equal(sanitizeMediaTitle(url), null);
  assert.equal(sanitizeMediaTitle('stream?static=true&api_key=test-secret'), null);
  assert.equal(sanitizeMediaTitle('stream static true api key test secret'), null);
  assert.equal(sanitizeMediaTitle('stream%3Fapi_key%3Dtest-secret'), null);
  assert.equal(sanitizeMediaTitle('  My Anime S01E02  '), 'My Anime S01E02');
  assert.equal(toMediaIdentityPath(url), 'jellyfin://jellyfin.example/item/item-1');
  assert.equal(
    toMediaIdentityPath('https://example.com/base/Videos/item-1/master.m3u8?token=secret'),
    'jellyfin://example.com/item/item-1',
  );
  assert.equal(toMediaIdentityPath('stream?api_key=test-secret'), '');
  assert.equal(toMediaIdentityPath('/media/My Anime S01E02.mkv'), '/media/My Anime S01E02.mkv');
});

test('remote stats identities drop credentials while preserving YouTube video identity', () => {
  assert.match(
    toMediaIdentityPath('https://user:password@example.com/video.mkv?signature=secret#secret'),
    /^https:\/\/example\.com\/video\.mkv#query-[a-f0-9]{64}$/,
  );
  assert.notEqual(
    toMediaIdentityPath('https://example.com/video?id=1'),
    toMediaIdentityPath('https://example.com/video?id=2'),
  );
  const identity = toMediaIdentityPath('https://example.com/video?id=1');
  assert.equal(toMediaIdentityPath(identity), identity);
  assert.equal(
    toMediaIdentityPath('https://www.youtube.com/watch?v=video-id&token=secret'),
    'https://www.youtube.com/watch?v=video-id',
  );
});
