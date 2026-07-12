import assert from 'node:assert/strict';
import test from 'node:test';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { execFileSync } from 'node:child_process';

import {
  animetoshoLangToFilenameSuffix,
  buildAnimetoshoAttachmentUrl,
  decompressXzFile,
  extractAnimetoshoSubtitleFiles,
  mapAnimetoshoSearchResults,
} from './utils.js';

test('buildAnimetoshoAttachmentUrl pads the attachment id to 8 hex digits', () => {
  assert.equal(
    buildAnimetoshoAttachmentUrl(1955356),
    'https://animetosho.org/storage/attach/001dd61c/1955356.xz',
  );
});

test('animetoshoLangToFilenameSuffix maps common ISO 639-2 codes to two-letter suffixes', () => {
  assert.equal(animetoshoLangToFilenameSuffix('eng'), 'en');
  assert.equal(animetoshoLangToFilenameSuffix('jpn'), 'ja');
  assert.equal(animetoshoLangToFilenameSuffix('ger'), 'de');
  assert.equal(animetoshoLangToFilenameSuffix('spa'), 'es');
  assert.equal(animetoshoLangToFilenameSuffix('POR'), 'pt');
});

test('animetoshoLangToFilenameSuffix falls back to the raw code, and to en when unknown', () => {
  assert.equal(animetoshoLangToFilenameSuffix('vie'), 'vie');
  assert.equal(animetoshoLangToFilenameSuffix(''), 'en');
  assert.equal(animetoshoLangToFilenameSuffix(undefined), 'en');
  assert.equal(animetoshoLangToFilenameSuffix('und'), 'en');
});

test('buildAnimetoshoAttachmentUrl rejects non-positive and non-integer ids', () => {
  assert.equal(buildAnimetoshoAttachmentUrl(0), null);
  assert.equal(buildAnimetoshoAttachmentUrl(-5), null);
  assert.equal(buildAnimetoshoAttachmentUrl(1.5), null);
  assert.equal(buildAnimetoshoAttachmentUrl(Number.NaN), null);
});

test('mapAnimetoshoSearchResults maps valid entries and caps to maxResults', () => {
  const payload = [
    {
      id: 606713,
      title: '[SubsPlease] Sousou no Frieren - 28 (1080p) [8BBBC28C].mkv',
      timestamp: 1710000000,
      total_size: 1490354395,
      num_files: 1,
    },
    { id: 'bogus', title: 'missing numeric id' },
    { id: 606714, title: '[Erai-raws] Sousou no Frieren - 28 [1080p].mkv' },
    { id: 606715, title: 'capped away' },
  ];

  const entries = mapAnimetoshoSearchResults(payload, 2);
  assert.equal(entries.length, 2);
  assert.deepEqual(entries[0], {
    id: 606713,
    title: '[SubsPlease] Sousou no Frieren - 28 (1080p) [8BBBC28C].mkv',
    timestamp: 1710000000,
    totalSize: 1490354395,
    numFiles: 1,
  });
  assert.equal(entries[1]!.id, 606714);
  assert.equal(entries[1]!.totalSize, null);
  assert.equal(entries[1]!.numFiles, null);
});

test('mapAnimetoshoSearchResults returns empty list for non-array payloads', () => {
  assert.deepEqual(mapAnimetoshoSearchResults({ error: 'nope' }, 10), []);
  assert.deepEqual(mapAnimetoshoSearchResults(null, 10), []);
});

const DETAIL_PAYLOAD = {
  id: 606713,
  title: '[SubsPlease] Sousou no Frieren - 28 (1080p) [8BBBC28C].mkv',
  files: [
    {
      id: 1151711,
      filename: '[SubsPlease] Sousou no Frieren - 28 (1080p) [8BBBC28C].mkv',
      attachments: [
        {
          id: 1955355,
          type: 'font',
          info: { name: 'arial.ttf' },
          size: 300000,
        },
        {
          id: 1955356,
          type: 'subtitle',
          info: { codec: 'ASS', lang: 'eng', name: 'English subs', trackid: 2 },
          size: 33075,
        },
      ],
    },
  ],
};

test('extractAnimetoshoSubtitleFiles keeps only text subtitle attachments with download urls', () => {
  const files = extractAnimetoshoSubtitleFiles(DETAIL_PAYLOAD);
  assert.equal(files.length, 1);
  const file = files[0]!;
  assert.equal(file.attachmentId, 1955356);
  assert.equal(file.lang, 'eng');
  assert.equal(file.trackName, 'English subs');
  assert.equal(file.size, 33075);
  assert.equal(file.url, 'https://animetosho.org/storage/attach/001dd61c/1955356.xz');
  assert.equal(file.sourceFilename, '[SubsPlease] Sousou no Frieren - 28 (1080p) [8BBBC28C].mkv');
  assert.equal(file.filename, '[SubsPlease] Sousou no Frieren - 28 (1080p) [8BBBC28C].eng.ass');
});

test('extractAnimetoshoSubtitleFiles skips image-based subtitle codecs', () => {
  const files = extractAnimetoshoSubtitleFiles({
    files: [
      {
        id: 1,
        filename: 'movie.mkv',
        attachments: [
          { id: 10, type: 'subtitle', info: { codec: 'PGS', lang: 'eng' }, size: 100 },
          { id: 11, type: 'subtitle', info: { codec: 'VobSub', lang: 'eng' }, size: 100 },
          { id: 12, type: 'subtitle', info: { codec: 'SRT', lang: 'eng' }, size: 100 },
        ],
      },
    ],
  });
  assert.deepEqual(
    files.map((f) => f.attachmentId),
    [12],
  );
  assert.equal(files[0]!.filename, 'movie.eng.srt');
});

test('extractAnimetoshoSubtitleFiles sorts English tracks first and disambiguates duplicates', () => {
  const files = extractAnimetoshoSubtitleFiles({
    files: [
      {
        id: 1,
        filename: 'episode.mkv',
        attachments: [
          {
            id: 21,
            type: 'subtitle',
            info: { codec: 'ASS', lang: 'ger', name: 'Deutsch' },
            size: 1,
          },
          {
            id: 22,
            type: 'subtitle',
            info: { codec: 'ASS', lang: 'eng', name: 'Signs & Songs' },
            size: 2,
          },
          {
            id: 23,
            type: 'subtitle',
            info: { codec: 'ASS', lang: 'eng', name: 'Full Subtitles' },
            size: 3,
          },
        ],
      },
    ],
  });

  assert.deepEqual(
    files.map((f) => f.attachmentId),
    [22, 23, 21],
  );
  assert.equal(files[0]!.filename, 'episode.eng.signs-songs.ass');
  assert.equal(files[1]!.filename, 'episode.eng.full-subtitles.ass');
  assert.equal(files[2]!.filename, 'episode.ger.ass');
});

test('extractAnimetoshoSubtitleFiles tolerates missing info fields', () => {
  const files = extractAnimetoshoSubtitleFiles({
    files: [
      {
        id: 1,
        attachments: [{ id: 31, type: 'subtitle', info: { codec: 'ASS' }, size: 5 }],
      },
    ],
  });
  assert.equal(files.length, 1);
  assert.equal(files[0]!.lang, '');
  assert.equal(files[0]!.trackName, null);
  assert.equal(files[0]!.filename, 'subtitle.ass');
});

const hasXz = (() => {
  try {
    execFileSync('xz', ['--version'], { stdio: 'ignore' });
    return true;
  } catch {
    return false;
  }
})();

test('decompressXzFile round-trips an xz-compressed subtitle', { skip: !hasXz }, async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-animetosho-test-'));
  try {
    const plainPath = path.join(dir, 'sub.ass');
    const content = '[Script Info]\nTitle: test\n';
    fs.writeFileSync(plainPath, content, 'utf8');
    execFileSync('xz', ['-z', plainPath]);

    const destPath = path.join(dir, 'out.ass');
    const result = await decompressXzFile(`${plainPath}.xz`, destPath);
    assert.equal(result.ok, true);
    if (result.ok) {
      assert.equal(result.path, destPath);
    }
    assert.equal(fs.readFileSync(destPath, 'utf8'), content);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('decompressXzFile reports an error for corrupt input', { skip: !hasXz }, async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-animetosho-test-'));
  try {
    const srcPath = path.join(dir, 'broken.xz');
    fs.writeFileSync(srcPath, 'not xz data');
    const result = await decompressXzFile(srcPath, path.join(dir, 'out.ass'));
    assert.equal(result.ok, false);
    if (!result.ok) {
      assert.match(result.error.error, /xz|decompress/i);
    }
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
