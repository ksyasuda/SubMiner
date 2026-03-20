import assert from 'node:assert/strict';
import test from 'node:test';
import {
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
