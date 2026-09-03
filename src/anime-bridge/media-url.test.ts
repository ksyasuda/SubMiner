import test from 'node:test';
import assert from 'node:assert/strict';
import { parseAnimeStatus, resolveBridgeMediaUrl, routeHlsThroughProxy } from './media-url';

const BRIDGE = 'http://127.0.0.1:56037';
const PROXY = 'http://127.0.0.1:60001';

test('a bridge m3u8 stream is routed through the strip proxy', () => {
  assert.equal(
    routeHlsThroughProxy(`${BRIDGE}/video/master.m3u8?q=1080`, BRIDGE, PROXY),
    `${PROXY}/video/master.m3u8?q=1080`,
  );
});

test('non-HLS bridge streams stay on the bridge', () => {
  const direct = `${BRIDGE}/video/movie-token`;
  assert.equal(routeHlsThroughProxy(direct, BRIDGE, PROXY), direct);
});

test('external m3u8 streams are not routed through the proxy', () => {
  const remote = 'https://cdn.example.com/hls/master.m3u8';
  assert.equal(routeHlsThroughProxy(remote, BRIDGE, PROXY), remote);
});

test('unparseable stream urls pass through routeHlsThroughProxy unchanged', () => {
  assert.equal(routeHlsThroughProxy('not a url', BRIDGE, PROXY), 'not a url');
  assert.equal(
    routeHlsThroughProxy(`${BRIDGE}/video/a.m3u8`, 'garbage', PROXY),
    `${BRIDGE}/video/a.m3u8`,
  );
});

test('a loopback proxy url is rebased onto the live bridge port', () => {
  assert.equal(
    resolveBridgeMediaUrl(BRIDGE, 'http://127.0.0.1:8080/image/cover-uuid'),
    'http://127.0.0.1:56037/image/cover-uuid',
  );
  assert.equal(
    resolveBridgeMediaUrl(BRIDGE, 'http://localhost:8080/video/master-token'),
    'http://127.0.0.1:56037/video/master-token',
  );
});

test('query strings and fragments survive rebasing', () => {
  assert.equal(
    resolveBridgeMediaUrl(BRIDGE, 'http://127.0.0.1:8080/video/token?quality=1080#t=30'),
    'http://127.0.0.1:56037/video/token?quality=1080#t=30',
  );
});

test('ipv6 loopback is recognised', () => {
  assert.equal(
    resolveBridgeMediaUrl(BRIDGE, 'http://[::1]:8080/image/cover'),
    'http://127.0.0.1:56037/image/cover',
  );
});

test('remote urls are left untouched', () => {
  const remote = 'https://cdn.example.com/covers/1.jpg';
  assert.equal(resolveBridgeMediaUrl(BRIDGE, remote), remote);
});

test('loopback urls outside the proxy routes are left untouched', () => {
  // Only /image and /video are proxy routes; /capabilities is the server's own API.
  const other = 'http://127.0.0.1:8080/capabilities';
  assert.equal(resolveBridgeMediaUrl(BRIDGE, other), other);
});

test('a base url without a scheme is assumed to be http', () => {
  assert.equal(
    resolveBridgeMediaUrl('127.0.0.1:56037', 'http://127.0.0.1:8080/image/cover'),
    'http://127.0.0.1:56037/image/cover',
  );
});

test('unparseable input is returned unchanged rather than throwing', () => {
  assert.equal(resolveBridgeMediaUrl(BRIDGE, 'not a url'), 'not a url');
  assert.equal(
    resolveBridgeMediaUrl('', 'http://127.0.0.1:8080/image/c'),
    'http://127.0.0.1:8080/image/c',
  );
  assert.equal(resolveBridgeMediaUrl(BRIDGE, ''), '');
});

test('parseAnimeStatus maps the SAnime constants', () => {
  assert.equal(parseAnimeStatus(1), 'ongoing');
  assert.equal(parseAnimeStatus(2), 'completed');
  assert.equal(parseAnimeStatus(4), 'publishing-finished');
  assert.equal(parseAnimeStatus(5), 'cancelled');
  assert.equal(parseAnimeStatus(6), 'on-hiatus');
  assert.equal(parseAnimeStatus(0), 'unknown');
  assert.equal(parseAnimeStatus(undefined), 'unknown');
  // 3 is unused in the SAnime constants.
  assert.equal(parseAnimeStatus(3), 'unknown');
});
