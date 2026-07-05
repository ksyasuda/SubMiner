import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Database } from 'bun:sqlite';
import {
  detectImageExtension,
  findNextEpisode,
  groupHistoryBySeries,
  isReadonlyWalRetryError,
  listSeasonDirs,
  materializeCoverArt,
  queryLocalWatchHistory,
  resolveSeriesRoot,
  seasonNumberFromDirName,
  sortVideosByEpisode,
  type HistoryVideoRow,
} from './history.js';

function makeRow(overrides: Partial<HistoryVideoRow> = {}): HistoryVideoRow {
  return {
    videoId: 1,
    sourcePath: '/media/anime/Show/Season-1/Show - S01E01.mkv',
    parsedTitle: 'Show',
    parsedSeason: 1,
    parsedEpisode: 1,
    animeTitle: null,
    lastWatchedMs: 1000,
    coverBlobHash: null,
    ...overrides,
  };
}

test('seasonNumberFromDirName detects common season directory names', () => {
  assert.equal(seasonNumberFromDirName('Season-1'), 1);
  assert.equal(seasonNumberFromDirName('Season 2'), 2);
  assert.equal(seasonNumberFromDirName('S03'), 3);
  assert.equal(seasonNumberFromDirName('season_04'), 4);
  assert.equal(seasonNumberFromDirName('Specials'), null);
  assert.equal(seasonNumberFromDirName('Show Name'), null);
});

test('resolveSeriesRoot skips season directories', () => {
  assert.equal(
    resolveSeriesRoot('/media/anime/Show/Season-1/Show - S01E01.mkv'),
    '/media/anime/Show',
  );
  assert.equal(resolveSeriesRoot('/media/anime/Show/Show - 01.mkv'), '/media/anime/Show');
});

test('groupHistoryBySeries keeps most recent entry per series root', () => {
  const rows = [
    makeRow({ videoId: 1, parsedEpisode: 1, lastWatchedMs: 1000 }),
    makeRow({
      videoId: 2,
      sourcePath: '/media/anime/Show/Season-1/Show - S01E02.mkv',
      parsedEpisode: 2,
      lastWatchedMs: 3000,
    }),
    makeRow({
      videoId: 3,
      sourcePath: '/media/anime/Other/Other - 05.mkv',
      parsedTitle: 'Other',
      parsedSeason: null,
      parsedEpisode: 5,
      lastWatchedMs: 2000,
    }),
  ];

  const series = groupHistoryBySeries(rows, () => true);

  assert.equal(series.length, 2);
  assert.equal(series[0]?.displayName, 'Show');
  assert.equal(series[0]?.seriesRoot, '/media/anime/Show');
  assert.equal(series[0]?.lastWatched.parsedEpisode, 2);
  assert.equal(series[1]?.displayName, 'Other');
});

test('groupHistoryBySeries filters series roots that no longer exist', () => {
  const rows = [
    makeRow({ videoId: 1 }),
    makeRow({
      videoId: 2,
      sourcePath: '/gone/anime/Missing/Season-1/Missing - S01E01.mkv',
      parsedTitle: 'Missing',
      lastWatchedMs: 5000,
    }),
  ];

  const series = groupHistoryBySeries(rows, (candidate) => !candidate.startsWith('/gone/'));

  assert.equal(series.length, 1);
  assert.equal(series[0]?.displayName, 'Show');
});

test('groupHistoryBySeries falls back to directory name for display', () => {
  const rows = [
    makeRow({
      sourcePath: '/media/anime/Some Show Dir/video.mkv',
      parsedTitle: null,
      animeTitle: null,
    }),
  ];

  const series = groupHistoryBySeries(rows, () => true);

  assert.equal(series[0]?.displayName, 'Some Show Dir');
});

test('sortVideosByEpisode orders by parsed episode with natural fallback', () => {
  const videos = [
    '/media/Show/Show - S01E10 - Ten.mkv',
    '/media/Show/Show - S01E02 - Two.mkv',
    '/media/Show/Show - S01E01 - One.mkv',
  ];

  assert.deepEqual(sortVideosByEpisode(videos), [
    '/media/Show/Show - S01E01 - One.mkv',
    '/media/Show/Show - S01E02 - Two.mkv',
    '/media/Show/Show - S01E10 - Ten.mkv',
  ]);
});

function createSeriesTree(): string {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-history-'));
  const seriesRoot = path.join(root, 'Show');
  const season1 = path.join(seriesRoot, 'Season-1');
  const season2 = path.join(seriesRoot, 'Season-2');
  fs.mkdirSync(season1, { recursive: true });
  fs.mkdirSync(season2, { recursive: true });
  fs.mkdirSync(path.join(seriesRoot, 'extras-empty'), { recursive: true });
  for (const name of ['Show - S01E01.mkv', 'Show - S01E02.mkv', 'Show - S01E03.mkv']) {
    fs.writeFileSync(path.join(season1, name), '');
  }
  fs.writeFileSync(path.join(season2, 'Show - S02E01.mkv'), '');
  fs.writeFileSync(path.join(season1, 'notes.txt'), '');
  return seriesRoot;
}

test('listSeasonDirs returns only video-bearing directories in season order', () => {
  const seriesRoot = createSeriesTree();
  try {
    const seasons = listSeasonDirs(seriesRoot);
    assert.deepEqual(
      seasons.map((entry) => entry.name),
      ['Season-1', 'Season-2'],
    );
    assert.deepEqual(
      seasons.map((entry) => entry.season),
      [1, 2],
    );
  } finally {
    fs.rmSync(path.dirname(seriesRoot), { recursive: true, force: true });
  }
});

test('findNextEpisode advances within a season and across seasons', () => {
  const seriesRoot = createSeriesTree();
  try {
    const season1 = path.join(seriesRoot, 'Season-1');
    const season2 = path.join(seriesRoot, 'Season-2');

    assert.equal(
      findNextEpisode(path.join(season1, 'Show - S01E02.mkv')),
      path.join(season1, 'Show - S01E03.mkv'),
    );
    assert.equal(
      findNextEpisode(path.join(season1, 'Show - S01E03.mkv')),
      path.join(season2, 'Show - S02E01.mkv'),
    );
    assert.equal(findNextEpisode(path.join(season2, 'Show - S02E01.mkv')), null);
  } finally {
    fs.rmSync(path.dirname(seriesRoot), { recursive: true, force: true });
  }
});

test('findNextEpisode falls back to episode numbers when file was removed', () => {
  const seriesRoot = createSeriesTree();
  try {
    const season1 = path.join(seriesRoot, 'Season-1');
    const missing = path.join(season1, 'Show - S01E02 - Deleted Cut.mkv');
    assert.equal(findNextEpisode(missing), path.join(season1, 'Show - S01E03.mkv'));
  } finally {
    fs.rmSync(path.dirname(seriesRoot), { recursive: true, force: true });
  }
});

test('findNextEpisode advances seasons when a deleted file was the last episode', () => {
  const seriesRoot = createSeriesTree();
  try {
    const season1 = path.join(seriesRoot, 'Season-1');
    const season2 = path.join(seriesRoot, 'Season-2');
    const missing = path.join(season1, 'Show - S01E03 - Deleted Cut.mkv');
    assert.equal(findNextEpisode(missing), path.join(season2, 'Show - S02E01.mkv'));
  } finally {
    fs.rmSync(path.dirname(seriesRoot), { recursive: true, force: true });
  }
});

const PNG_MAGIC = Buffer.from('89504e470d0a1a0a0000000d49484452', 'hex');

function createHistoryDb(
  dbPath: string,
  options: { wal?: boolean; coverArt?: boolean } = {},
): void {
  const db = new Database(dbPath);
  try {
    if (options.wal) db.run('PRAGMA journal_mode = WAL;');
    db.run(`
      CREATE TABLE imm_anime(
        anime_id INTEGER PRIMARY KEY,
        canonical_title TEXT,
        title_romaji TEXT
      );
    `);
    db.run(`
      CREATE TABLE imm_videos(
        video_id INTEGER PRIMARY KEY,
        anime_id INTEGER,
        source_type INTEGER,
        source_path TEXT,
        parsed_title TEXT,
        parsed_season INTEGER,
        parsed_episode INTEGER
      );
    `);
    db.run(`
      CREATE TABLE imm_sessions(
        session_id INTEGER PRIMARY KEY,
        video_id INTEGER,
        started_at_ms TEXT
      );
    `);
    db.run(`INSERT INTO imm_anime VALUES (1, 'Show Season 1', 'Show Romaji');`);
    db.run(`
      INSERT INTO imm_videos VALUES
        (1, 1, 1, '/media/Show/Season-1/Show - S01E01.mkv', 'Show', 1, 1),
        (2, 1, 1, '/media/Show/Season-1/Show - S01E02.mkv', 'Show', 1, 2),
        (3, NULL, 2, NULL, 'Remote Show', NULL, NULL),
        (4, NULL, 1, '', 'Empty Path', NULL, NULL);
    `);
    db.run(`
      INSERT INTO imm_sessions VALUES
        (1, 1, '1000'),
        (2, 1, '5000'),
        (3, 2, '3000'),
        (4, 3, '9000');
    `);
    if (options.coverArt) {
      db.run(`
        CREATE TABLE imm_media_art(
          video_id INTEGER PRIMARY KEY,
          cover_blob_hash TEXT
        );
      `);
      db.run(`
        CREATE TABLE imm_cover_art_blobs(
          blob_hash TEXT PRIMARY KEY,
          cover_blob BLOB NOT NULL
        );
      `);
      // Art only on video 1; video 2 resolves it through the shared anime_id.
      db.run(`INSERT INTO imm_media_art VALUES (1, 'hash-1');`);
      db.query('INSERT INTO imm_cover_art_blobs VALUES (?, ?)').run('hash-1', PNG_MAGIC);
    }
    if (options.wal) db.run('PRAGMA wal_checkpoint(TRUNCATE);');
  } finally {
    db.close();
  }
}

function assertHistoryRows(dbPath: string): void {
  const rows = queryLocalWatchHistory(dbPath);

  assert.equal(rows.length, 2);
  assert.equal(rows[0]?.videoId, 1);
  assert.equal(rows[0]?.lastWatchedMs, 5000);
  assert.equal(rows[0]?.animeTitle, 'Show Romaji');
  assert.equal(rows[1]?.videoId, 2);
  assert.equal(rows[1]?.lastWatchedMs, 3000);
}

test('queryLocalWatchHistory returns local files ordered by most recent session', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-history-db-'));
  const dbPath = path.join(dir, 'immersion.sqlite');
  try {
    createHistoryDb(dbPath);
    assertHistoryRows(dbPath);
    const rows = queryLocalWatchHistory(dbPath);
    assert.equal(rows[0]?.coverBlobHash, null);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('queryLocalWatchHistory resolves cover hashes directly and via shared anime', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-history-art-'));
  const dbPath = path.join(dir, 'immersion.sqlite');
  try {
    createHistoryDb(dbPath, { coverArt: true });
    const rows = queryLocalWatchHistory(dbPath);
    assert.equal(rows[0]?.videoId, 1);
    assert.equal(rows[0]?.coverBlobHash, 'hash-1');
    assert.equal(rows[1]?.videoId, 2);
    assert.equal(rows[1]?.coverBlobHash, 'hash-1');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('materializeCoverArt extracts blobs to the cache dir and reuses cached files', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-history-covers-'));
  const dbPath = path.join(dir, 'immersion.sqlite');
  const cacheDir = path.join(dir, 'covers');
  try {
    createHistoryDb(dbPath, { coverArt: true });

    const covers = materializeCoverArt(
      dbPath,
      ['hash-1', 'hash-1', null, 'hash-missing'],
      cacheDir,
    );
    const coverPath = covers.get('hash-1');
    assert.ok(coverPath);
    assert.equal(path.extname(coverPath!), '.png');
    assert.ok(fs.statSync(coverPath!).size > 0);
    assert.equal(covers.has('hash-missing'), false);

    // Cached file is reused even when the database has disappeared.
    fs.rmSync(dbPath);
    const cachedCovers = materializeCoverArt(dbPath, ['hash-1'], cacheDir);
    assert.equal(cachedCovers.get('hash-1'), coverPath);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('materializeCoverArt rejects cover hashes that escape the cache dir', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-history-cover-safety-'));
  const dbPath = path.join(dir, 'immersion.sqlite');
  const cacheDir = path.join(dir, 'covers');
  const unsafeHash = '../escape';
  try {
    createHistoryDb(dbPath, { coverArt: true });
    const db = new Database(dbPath);
    try {
      db.query('INSERT INTO imm_cover_art_blobs VALUES (?, ?)').run(unsafeHash, PNG_MAGIC);
    } finally {
      db.close();
    }

    const covers = materializeCoverArt(dbPath, [unsafeHash], cacheDir);

    assert.equal(covers.has(unsafeHash), false);
    assert.equal(fs.existsSync(path.join(dir, 'escape.png')), false);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('detectImageExtension identifies common cover formats', () => {
  assert.equal(detectImageExtension(PNG_MAGIC), '.png');
  assert.equal(detectImageExtension(Buffer.from([0xff, 0xd8, 0xff, 0xe0])), '.jpg');
  assert.equal(detectImageExtension(Buffer.from('RIFF0000WEBPVP8 ', 'ascii')), '.webp');
  assert.equal(detectImageExtension(Buffer.from('GIF89a', 'ascii')), '.gif');
  assert.equal(detectImageExtension(Buffer.from('unknown', 'ascii')), '.jpg');
});

test('groupHistoryBySeries backfills cover hash from older rows of the same series', () => {
  const rows = [
    makeRow({ videoId: 2, parsedEpisode: 2, lastWatchedMs: 3000, coverBlobHash: null }),
    makeRow({ videoId: 1, parsedEpisode: 1, lastWatchedMs: 1000, coverBlobHash: 'hash-1' }),
  ];

  const series = groupHistoryBySeries(rows, () => true);

  assert.equal(series.length, 1);
  assert.equal(series[0]?.lastWatched.videoId, 2);
  assert.equal(series[0]?.coverBlobHash, 'hash-1');
});

test('queryLocalWatchHistory reads a cleanly-closed WAL database', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-history-wal-'));
  const dbPath = path.join(dir, 'immersion.sqlite');
  try {
    createHistoryDb(dbPath, { wal: true });
    // Reproduce the state after the app shuts down cleanly: WAL journal mode
    // with no -wal/-shm sidecar files on disk. A read-only connection then
    // fails at query time because it cannot recreate them.
    fs.rmSync(`${dbPath}-wal`, { force: true });
    fs.rmSync(`${dbPath}-shm`, { force: true });
    assertHistoryRows(dbPath);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('isReadonlyWalRetryError only accepts readonly errors from WAL-mode databases', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-history-retry-'));
  const walDbPath = path.join(dir, 'wal.sqlite');
  const rollbackDbPath = path.join(dir, 'rollback.sqlite');
  try {
    createHistoryDb(walDbPath, { wal: true });
    createHistoryDb(rollbackDbPath);

    assert.equal(
      isReadonlyWalRetryError(
        Object.assign(new Error('attempt to write a readonly database'), {
          code: 'SQLITE_READONLY',
        }),
        walDbPath,
      ),
      true,
    );
    assert.equal(
      isReadonlyWalRetryError(
        Object.assign(new Error('unable to open database file'), {
          code: 'SQLITE_CANTOPEN',
        }),
        walDbPath,
      ),
      true,
    );
    assert.equal(
      isReadonlyWalRetryError(new Error('no such table: imm_sessions'), walDbPath),
      false,
    );
    assert.equal(
      isReadonlyWalRetryError(
        Object.assign(new Error('attempt to write a readonly database'), {
          code: 'SQLITE_READONLY',
        }),
        rollbackDbPath,
      ),
      false,
    );
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
