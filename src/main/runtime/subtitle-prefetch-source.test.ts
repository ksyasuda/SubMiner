import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildSubtitleSidebarSourceKey,
  getActiveExternalSubtitleSource,
  resolveSubtitleSourcePath,
} from './subtitle-prefetch-source';

test('getActiveExternalSubtitleSource returns the active external subtitle path', () => {
  const source = getActiveExternalSubtitleSource(
    [
      { type: 'sub', id: 1, external: false },
      { type: 'sub', id: 2, external: true, 'external-filename': ' https://host/subs.ass ' },
    ],
    '2',
  );

  assert.equal(source, 'https://host/subs.ass');
});

test('getActiveExternalSubtitleSource normalizes integer-like string track ids', () => {
  const source = getActiveExternalSubtitleSource(
    [{ type: 'sub', id: '2', external: true, 'external-filename': ' /tmp/subs.ass ' }],
    '2',
  );

  assert.equal(source, '/tmp/subs.ass');
});

test('getActiveExternalSubtitleSource returns null when the selected track is not external', () => {
  const source = getActiveExternalSubtitleSource(
    [{ type: 'sub', id: 2, external: false, 'external-filename': '/tmp/subs.ass' }],
    2,
  );

  assert.equal(source, null);
});

test('resolveSubtitleSourcePath converts file URLs with spaces into filesystem paths', () => {
  const fileUrl =
    process.platform === 'win32'
      ? 'file:///C:/Users/test/Sub%20Folder/subs.ass'
      : 'file:///tmp/Sub%20Folder/subs.ass';

  const resolved = resolveSubtitleSourcePath(fileUrl);

  assert.ok(
    resolved.endsWith('/Sub Folder/subs.ass') || resolved.endsWith('\\Sub Folder\\subs.ass'),
  );
});

test('resolveSubtitleSourcePath leaves non-file sources unchanged', () => {
  assert.equal(resolveSubtitleSourcePath('/tmp/subs.ass'), '/tmp/subs.ass');
});

test('resolveSubtitleSourcePath returns the original source for malformed file URLs', () => {
  const source = 'file://invalid[path';

  assert.equal(resolveSubtitleSourcePath(source), source);
});

test('buildSubtitleSidebarSourceKey uses a stable identifier for internal subtitle tracks', () => {
  const firstKey = buildSubtitleSidebarSourceKey('/media/episode01.mkv', {
    id: 3,
    'ff-index': 7,
    title: 'English',
    lang: 'eng',
    codec: 'ass',
  });
  const secondKey = buildSubtitleSidebarSourceKey('/media/episode01.mkv', {
    id: 3,
    'ff-index': 7,
    title: 'English',
    lang: 'eng',
    codec: 'ass',
  });

  assert.equal(firstKey, secondKey);
  assert.equal(firstKey, 'internal:/media/episode01.mkv:track:3:ff:7');
});

test('buildSubtitleSidebarSourceKey normalizes integer-like string track metadata', () => {
  const key = buildSubtitleSidebarSourceKey('/media/episode01.mkv', {
    id: '3',
    'ff-index': '7',
  });

  assert.equal(key, 'internal:/media/episode01.mkv:track:3:ff:7');
});

test('buildSubtitleSidebarSourceKey falls back to source path when no track metadata is available', () => {
  const key = buildSubtitleSidebarSourceKey('/media/episode01.mkv', null, '/tmp/subtitle.ass');

  assert.equal(key, '/tmp/subtitle.ass');
});
