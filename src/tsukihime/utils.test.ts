import assert from 'node:assert/strict';
import test from 'node:test';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { execFileSync } from 'node:child_process';

import {
  tsukihimeLangToFilenameSuffix,
  buildTsukihimeAttachmentUrl,
  decompressXzFile,
  extractTsukihimeSubtitleFiles,
  isTsukihimeDownloadUrl,
  mapTsukihimeSearchResults,
} from './utils.js';

test('buildTsukihimeAttachmentUrl pads the attachment id to 8 hex digits', () => {
  assert.equal(
    buildTsukihimeAttachmentUrl(96412),
    'https://storage.tsukihime.org/attach/0001789c/96412.xz',
  );
});

test('buildTsukihimeAttachmentUrl uses the tosho mirror path for imported entries', () => {
  assert.equal(
    buildTsukihimeAttachmentUrl(2826332, { imported: true }),
    'https://storage.tsukihime.org/tosho/attach/002b205c/2826332.xz',
  );
});

test('tsukihimeLangToFilenameSuffix maps common ISO 639-2 codes to two-letter suffixes', () => {
  assert.equal(tsukihimeLangToFilenameSuffix('eng'), 'en');
  assert.equal(tsukihimeLangToFilenameSuffix('jpn'), 'ja');
  assert.equal(tsukihimeLangToFilenameSuffix('ger'), 'de');
  assert.equal(tsukihimeLangToFilenameSuffix('spa'), 'es');
  assert.equal(tsukihimeLangToFilenameSuffix('POR'), 'pt');
});

test('tsukihimeLangToFilenameSuffix handles TsukiHime BCP-47 style codes', () => {
  assert.equal(tsukihimeLangToFilenameSuffix('en-US'), 'en');
  assert.equal(tsukihimeLangToFilenameSuffix('es-419'), 'es');
  assert.equal(tsukihimeLangToFilenameSuffix('zh-Hant'), 'zh');
  assert.equal(tsukihimeLangToFilenameSuffix('ja'), 'ja');
});

test('tsukihimeLangToFilenameSuffix falls back to the raw code, and to en when unknown', () => {
  assert.equal(tsukihimeLangToFilenameSuffix('vie'), 'vie');
  assert.equal(tsukihimeLangToFilenameSuffix(''), 'en');
  assert.equal(tsukihimeLangToFilenameSuffix(undefined), 'en');
  assert.equal(tsukihimeLangToFilenameSuffix('und'), 'en');
});

test('buildTsukihimeAttachmentUrl rejects non-positive and non-integer ids', () => {
  assert.equal(buildTsukihimeAttachmentUrl(0), null);
  assert.equal(buildTsukihimeAttachmentUrl(-5), null);
  assert.equal(buildTsukihimeAttachmentUrl(1.5), null);
  assert.equal(buildTsukihimeAttachmentUrl(Number.NaN), null);
});

test('isTsukihimeDownloadUrl accepts tsukihime.org and animetosho.org hosts only', () => {
  assert.equal(isTsukihimeDownloadUrl('https://storage.tsukihime.org/attach/0001789c/1.xz'), true);
  assert.equal(isTsukihimeDownloadUrl('https://tsukihime.org/view/1'), true);
  // The tosho mirror currently 302s to storage.animetosho.org; the hop must stay allowed.
  assert.equal(
    isTsukihimeDownloadUrl('https://storage.animetosho.org/attach/002b205c/2826332.xz'),
    true,
  );
  assert.equal(isTsukihimeDownloadUrl('http://storage.tsukihime.org/attach/1/1.xz'), false);
  assert.equal(isTsukihimeDownloadUrl('https://eviltsukihime.org/attach/1/1.xz'), false);
  assert.equal(isTsukihimeDownloadUrl('https://example.com/1.xz'), false);
  assert.equal(isTsukihimeDownloadUrl('not a url'), false);
});

test('mapTsukihimeSearchResults maps valid results and caps to maxResults', () => {
  const payload = {
    total: 4,
    start: 0,
    limit: 50,
    error: false,
    results: [
      {
        id: 1000752022,
        state: 'completed',
        name: '[SubsPlease] Sousou no Frieren S2 - 10 (1080p) [7D35515E].mkv',
        totalsize: 1457009569,
        filecount: 1,
        sublangs: ['en'],
        audiolangs: ['ja'],
        source_date: 1774623761,
        added_date: 1774624861,
        animetosho: true,
      },
      { id: 'bogus', name: 'missing numeric id' },
      { id: 12255, name: '[DKB] Futsutsuka na Akujo - S01E01 [Multi-Subs]' },
      { id: 12256, name: 'capped away' },
    ],
  };

  const entries = mapTsukihimeSearchResults(payload, 2);
  assert.equal(entries.length, 2);
  assert.deepEqual(entries[0], {
    id: 1000752022,
    title: '[SubsPlease] Sousou no Frieren S2 - 10 (1080p) [7D35515E].mkv',
    timestamp: 1774624861,
    totalSize: 1457009569,
    numFiles: 1,
    sublangs: ['en'],
  });
  assert.equal(entries[1]!.id, 12255);
  assert.equal(entries[1]!.totalSize, null);
  assert.equal(entries[1]!.numFiles, null);
  assert.deepEqual(entries[1]!.sublangs, []);
});

test('mapTsukihimeSearchResults returns empty list for malformed payloads', () => {
  assert.deepEqual(mapTsukihimeSearchResults({ error: 'nope' }, 10), []);
  assert.deepEqual(mapTsukihimeSearchResults({ results: 'nope' }, 10), []);
  assert.deepEqual(mapTsukihimeSearchResults(null, 10), []);
  assert.deepEqual(mapTsukihimeSearchResults([], 10), []);
});

// Trimmed from a real /v1/torrents/{id} response (native entry).
const DETAIL_PAYLOAD = {
  id: 12255,
  name: '[DKB] Futsutsuka na Akujo dewa Gozaimasu ga - S01E01 [Multi-Subs]',
  files: [
    {
      id: 16179,
      filename: '[DKB] Futsutsuka na Akujo dewa Gozaimasu ga - S01E01 [Multi-Subs].mkv',
      attachments: [
        {
          id: 96400,
          type: 0,
          info: { name: 'arial.ttf', mime: 'application/x-truetype-font' },
        },
        {
          id: 96412,
          type: 1,
          info: { codec: 'ASS', lang: 'en-US', name: 'English subs', trackid: 2, tracknum: 3 },
        },
        { id: 96499, type: 3 },
      ],
    },
  ],
};

test('extractTsukihimeSubtitleFiles keeps only text subtitle attachments with download urls', () => {
  const files = extractTsukihimeSubtitleFiles(DETAIL_PAYLOAD);
  assert.equal(files.length, 1);
  const file = files[0]!;
  assert.equal(file.attachmentId, 96412);
  assert.equal(file.lang, 'en-us');
  assert.equal(file.trackName, 'English subs');
  assert.equal(file.size, 0);
  assert.equal(file.url, 'https://storage.tsukihime.org/attach/0001789c/96412.xz');
  assert.equal(
    file.sourceFilename,
    '[DKB] Futsutsuka na Akujo dewa Gozaimasu ga - S01E01 [Multi-Subs].mkv',
  );
  assert.equal(
    file.filename,
    '[DKB] Futsutsuka na Akujo dewa Gozaimasu ga - S01E01 [Multi-Subs].en-us.ass',
  );
});

test('extractTsukihimeSubtitleFiles uses the tosho mirror for animetosho-imported entries', () => {
  const files = extractTsukihimeSubtitleFiles({
    id: 1000752022,
    animetosho: true,
    files: [
      {
        id: 1001361385,
        filename: '[SubsPlease] Sousou no Frieren S2 - 10 (1080p) [7D35515E].mkv',
        attachments: [
          { id: 2826332, type: 1, info: { codec: 'ASS', lang: 'en', name: 'English subs' } },
        ],
      },
    ],
  });
  assert.equal(files.length, 1);
  assert.equal(files[0]!.url, 'https://storage.tsukihime.org/tosho/attach/002b205c/2826332.xz');
});

test('extractTsukihimeSubtitleFiles detects imported entries by id when the flag is absent', () => {
  const files = extractTsukihimeSubtitleFiles({
    id: 1000752022,
    files: [
      {
        id: 1001361385,
        filename: 'episode.mkv',
        attachments: [{ id: 2826332, type: 1, info: { codec: 'ASS', lang: 'en' } }],
      },
    ],
  });
  assert.equal(files[0]!.url, 'https://storage.tsukihime.org/tosho/attach/002b205c/2826332.xz');
});

test('extractTsukihimeSubtitleFiles skips image-based subtitle codecs', () => {
  const files = extractTsukihimeSubtitleFiles({
    id: 1,
    files: [
      {
        id: 1,
        filename: 'movie.mkv',
        attachments: [
          { id: 10, type: 1, info: { codec: 'PGS', lang: 'en' } },
          { id: 11, type: 1, info: { codec: 'VobSub', lang: 'en' } },
          { id: 12, type: 1, info: { codec: 'SRT', lang: 'en' } },
        ],
      },
    ],
  });
  assert.deepEqual(
    files.map((f) => f.attachmentId),
    [12],
  );
  assert.equal(files[0]!.filename, 'movie.en.srt');
});

test('extractTsukihimeSubtitleFiles sorts English tracks first and disambiguates duplicates', () => {
  const files = extractTsukihimeSubtitleFiles({
    id: 1,
    files: [
      {
        id: 1,
        filename: 'episode.mkv',
        attachments: [
          { id: 21, type: 1, info: { codec: 'ASS', lang: 'de', name: 'Deutsch' } },
          { id: 22, type: 1, info: { codec: 'ASS', lang: 'en', name: 'Signs & Songs' } },
          { id: 23, type: 1, info: { codec: 'ASS', lang: 'en', name: 'Full Subtitles' } },
        ],
      },
    ],
  });

  assert.deepEqual(
    files.map((f) => f.attachmentId),
    [22, 23, 21],
  );
  assert.equal(files[0]!.filename, 'episode.en.signs-songs.ass');
  assert.equal(files[1]!.filename, 'episode.en.full-subtitles.ass');
  assert.equal(files[2]!.filename, 'episode.de.ass');
});

test('extractTsukihimeSubtitleFiles tolerates missing info fields', () => {
  const files = extractTsukihimeSubtitleFiles({
    id: 1,
    files: [
      {
        id: 1,
        attachments: [{ id: 31, type: 1, info: { codec: 'ASS' } }],
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
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-tsukihime-test-'));
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
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-tsukihime-test-'));
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
