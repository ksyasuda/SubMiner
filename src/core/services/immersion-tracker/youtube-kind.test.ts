import assert from 'node:assert/strict';
import test from 'node:test';
import { Database, type DatabaseSync } from './sqlite';
import {
  ensureSchema,
  getOrCreateAnimeRecord,
  getOrCreateVideoRecord,
  linkYoutubeVideoToAnimeRecord,
} from './storage';
import { getAnimeDetail, getAnimeLibrary } from './query-library';
import { updateAnimeAnilistInfo } from './query-maintenance';
import {
  repairLegacySeasonlessAnimeRows,
  resolveAnimeAnilistConflict,
} from './anime-season-repair';
import {
  getAnimeMergeRecommendations,
  recordAnimeMergeRecommendation,
} from './anime-merge-recommendations';
import { createCoverArtFetcher } from '../anilist/cover-art-fetcher';
import { SCHEMA_VERSION, SOURCE_TYPE_REMOTE, type YoutubeVideoMetadata } from './types';

const metadata: YoutubeVideoMetadata = {
  youtubeVideoId: 'video1',
  videoUrl: 'https://www.youtube.com/watch?v=video1',
  videoTitle: 'Video title',
  videoThumbnailUrl: null,
  channelId: 'UC123',
  channelName: 'Channel name',
  channelUrl: 'https://www.youtube.com/channel/UC123',
  channelThumbnailUrl: null,
  uploaderId: null,
  uploaderUrl: null,
  description: null,
  metadataJson: null,
};

function createVideo(db: DatabaseSync, key: string): number {
  return getOrCreateVideoRecord(db, key, {
    canonicalTitle: 'Video title',
    sourcePath: null,
    sourceUrl: metadata.videoUrl,
    sourceType: SOURCE_TYPE_REMOTE,
  });
}

function createAnime(
  db: DatabaseSync,
  parsedTitle: string,
  metadataJson: string | null = null,
): number {
  return getOrCreateAnimeRecord(db, {
    parsedTitle,
    canonicalTitle: parsedTitle,
    metadataJson,
    anilistId: null,
    titleRomaji: null,
    titleEnglish: null,
    titleNative: null,
  });
}

test('schema 23 channel migration preserves history and manual assignments and is idempotent', () => {
  const db = new Database(':memory:');
  try {
    ensureSchema(db);
    const ids = [
      createAnime(db, 'youtube-channel:UC123'),
      createAnime(db, 'youtube-channel-url:https://www.youtube.com/@creator'),
      createAnime(db, 'youtube-channel-name:Creator'),
      createAnime(db, 'Renamed channel', '{ "source": "youtube-channel" }'),
    ];
    const animeId = createAnime(db, 'Anime title', 'legacy non-JSON metadata');
    const videoId = createVideo(db, 'manual');
    db.prepare(
      'UPDATE imm_videos SET anime_id = ?, anime_assignment_locked = 1 WHERE video_id = ?',
    ).run(animeId, videoId);
    db.prepare(
      'INSERT INTO imm_lifetime_anime(anime_id, total_active_ms, total_cards) VALUES (?, 123456, 7)',
    ).run(animeId);
    const history = getAnimeLibrary(db);
    // Reproduce the previous schema, including its lack of a media kind column.
    db.exec(
      'DROP INDEX idx_anime_kind_title; ALTER TABLE imm_anime DROP COLUMN media_kind; DELETE FROM imm_schema_version; INSERT INTO imm_schema_version VALUES (23, 0)',
    );
    ensureSchema(db);
    ensureSchema(db);
    for (const id of ids) {
      const row = db.prepare('SELECT media_kind FROM imm_anime WHERE anime_id = ?').get(id);
      assert.ok(row && typeof row === 'object' && 'media_kind' in row);
      assert.equal(row.media_kind, 'youtube');
    }
    assert.deepEqual(getAnimeLibrary(db), history);
    assert.equal(linkYoutubeVideoToAnimeRecord(db, videoId, metadata), animeId);
    const version = db
      .prepare('SELECT MAX(schema_version) AS version FROM imm_schema_version')
      .get();
    assert.ok(version && typeof version === 'object' && 'version' in version);
    assert.equal(version.version, SCHEMA_VERSION);
  } finally {
    db.close();
  }
});

test('channel creation and repeated linking expose youtube in library and detail without losing totals', async () => {
  const db = new Database(':memory:');
  try {
    ensureSchema(db);
    const videoId = createVideo(db, 'first');
    const channelId = linkYoutubeVideoToAnimeRecord(db, videoId, metadata);
    assert.ok(channelId);
    const secondVideoId = createVideo(db, 'second');
    assert.equal(linkYoutubeVideoToAnimeRecord(db, secondVideoId, metadata), channelId);
    db.prepare(
      'INSERT INTO imm_lifetime_anime(anime_id, total_active_ms, total_cards) VALUES (?, 123456, 7)',
    ).run(channelId);
    assert.equal(getAnimeLibrary(db)[0]?.mediaKind, 'youtube');
    const detail = getAnimeDetail(db, channelId);
    assert.equal(detail?.mediaKind, 'youtube');
    assert.equal(detail?.episodeCount, 2);
    assert.equal(detail?.totalActiveMs, 123456);
    assert.equal(detail?.totalCards, 7);

    // Even parsed season numbers and a matching anime name must not trigger repairs.
    db.prepare('UPDATE imm_videos SET parsed_season = video_id').run();
    assert.equal(repairLegacySeasonlessAnimeRows(db).repaired, 0);
    const animeId = createAnime(db, 'Channel name');
    for (const matchConfidence of ['exact', 'weak', 'manual'] as const) {
      assert.equal(
        resolveAnimeAnilistConflict(db, channelId, 123, { matchConfidence })
          .anilistAssignmentBlocked,
        true,
      );
    }
    updateAnimeAnilistInfo(db, videoId, {
      anilistId: 123,
      titleRomaji: 'Wrong title',
      titleEnglish: null,
      titleNative: null,
      episodesTotal: 12,
    });
    recordAnimeMergeRecommendation(db, channelId, animeId, 123);
    assert.deepEqual(getAnimeMergeRecommendations(db), []);
    assert.equal(getAnimeDetail(db, channelId)?.anilistId, null);
    const fetcher = createCoverArtFetcher(
      {
        acquire: async () => {
          assert.fail('YouTube must not query AniList');
        },
        recordResponse: () => {},
      },
      console,
    );
    assert.equal(await fetcher.fetchIfMissing(db, videoId, 'Channel name'), false);
  } finally {
    db.close();
  }
});

test('startup reclassifies channels created by an older build after the schema upgrade', () => {
  const db = new Database(':memory:');
  try {
    ensureSchema(db);
    // An old build omits media_kind when creating a channel in the upgraded DB.
    const channelId = createAnime(db, 'youtube-channel:UCnew');
    db.prepare('UPDATE imm_anime SET anilist_id = 321 WHERE anime_id = ?').run(channelId);
    const animeId = createAnime(db, 'Regular anime');
    const videoId = createVideo(db, 'older-build');
    db.prepare(
      'UPDATE imm_videos SET anime_id = ?, anime_assignment_locked = 1 WHERE video_id = ?',
    ).run(channelId, videoId);
    db.prepare(
      'INSERT INTO imm_lifetime_anime(anime_id, total_active_ms, total_cards) VALUES (?, 120000, 5)',
    ).run(channelId);
    assert.equal(getAnimeLibrary(db)[0]?.mediaKind, 'anime');
    ensureSchema(db);
    const channel = getAnimeLibrary(db)[0];
    assert.equal(channel?.animeId, channelId);
    assert.equal(channel?.mediaKind, 'youtube');
    assert.equal(
      (
        db.prepare('SELECT anilist_id FROM imm_anime WHERE anime_id = ?').get(channelId) as {
          anilist_id: number | null;
        }
      ).anilist_id,
      null,
    );
    assert.equal(channel?.totalActiveMs, 120000);
    assert.equal(channel?.totalCards, 5);
    assert.equal(linkYoutubeVideoToAnimeRecord(db, videoId, metadata), channelId);
    const anime = db.prepare('SELECT media_kind FROM imm_anime WHERE anime_id = ?').get(animeId);
    assert.ok(anime && typeof anime === 'object' && 'media_kind' in anime);
    assert.equal(anime.media_kind, 'anime');
  } finally {
    db.close();
  }
});

test('title identity and aliases never cross media kinds', () => {
  const db = new Database(':memory:');
  try {
    ensureSchema(db);
    const animeId = createAnime(db, 'Shared title');
    const input = {
      mediaKind: 'youtube' as const,
      parsedTitle: 'Shared title',
      canonicalTitle: 'Shared title',
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: null,
    };
    const channelId = getOrCreateAnimeRecord(db, input);
    assert.notEqual(channelId, animeId);
    db.prepare('UPDATE imm_anime SET anilist_id = 123 WHERE anime_id = ?').run(channelId);
    assert.equal(getOrCreateAnimeRecord(db, input), channelId);
    assert.equal(
      (
        db.prepare('SELECT anilist_id FROM imm_anime WHERE anime_id = ?').get(channelId) as {
          anilist_id: number | null;
        }
      ).anilist_id,
      null,
    );
    assert.equal(createAnime(db, 'Shared title'), animeId);
    db.prepare(
      'INSERT INTO imm_anime_title_aliases(normalized_title_key, anime_id) VALUES (?, ?)',
    ).run('alias title', animeId);
    assert.notEqual(getOrCreateAnimeRecord(db, { ...input, parsedTitle: 'Alias title' }), animeId);
    assert.throws(() =>
      db
        .prepare(
          "INSERT INTO imm_anime(normalized_title_key, canonical_title, media_kind) VALUES ('shared title', 'duplicate', 'youtube')",
        )
        .run(),
    );
  } finally {
    db.close();
  }
});

test('schema 24 title constraint migration preserves referenced data', () => {
  const db = new Database(':memory:');
  try {
    // Reproduce the original table-level uniqueness constraint.
    db.exec(`CREATE TABLE imm_anime(
      anime_id INTEGER PRIMARY KEY AUTOINCREMENT,
      normalized_title_key TEXT NOT NULL UNIQUE,
      canonical_title TEXT NOT NULL,
      anilist_id INTEGER UNIQUE,
      title_romaji TEXT, title_english TEXT, title_native TEXT,
      episodes_total INTEGER, description TEXT, metadata_json TEXT,
      CREATED_DATE TEXT, LAST_UPDATE_DATE TEXT,
      media_kind TEXT NOT NULL DEFAULT 'anime' CHECK(media_kind IN ('anime', 'youtube'))
    )`);
    ensureSchema(db);
    const animeId = createAnime(db, 'Shared title');
    const videoId = createVideo(db, 'migration-video');
    db.prepare(
      'UPDATE imm_videos SET anime_id = ?, anime_assignment_locked = 1 WHERE video_id = ?',
    ).run(animeId, videoId);
    db.prepare(
      'INSERT INTO imm_lifetime_anime(anime_id, total_active_ms, total_cards) VALUES (?, 123, 4)',
    ).run(animeId);
    // Restore the old constraint while leaving child rows populated.
    db.exec(`PRAGMA foreign_keys = OFF;
      CREATE TABLE old_anime AS SELECT * FROM imm_anime;
      DROP TABLE imm_anime;
      CREATE TABLE imm_anime(
        anime_id INTEGER PRIMARY KEY AUTOINCREMENT, normalized_title_key TEXT NOT NULL UNIQUE,
        canonical_title TEXT NOT NULL, anilist_id INTEGER UNIQUE,
        title_romaji TEXT, title_english TEXT, title_native TEXT, episodes_total INTEGER,
        description TEXT, metadata_json TEXT, CREATED_DATE TEXT, LAST_UPDATE_DATE TEXT,
        media_kind TEXT NOT NULL DEFAULT 'anime' CHECK(media_kind IN ('anime', 'youtube')));
      INSERT INTO imm_anime SELECT * FROM old_anime;
      DROP TABLE old_anime;
      DELETE FROM imm_schema_version;
      INSERT INTO imm_schema_version VALUES (24, 0);
      PRAGMA foreign_keys = ON;`);
    ensureSchema(db);
    ensureSchema(db);
    assert.deepEqual(db.prepare('PRAGMA foreign_key_check').all(), []);
    assert.equal(
      (db.prepare('PRAGMA foreign_keys').get() as { foreign_keys: number }).foreign_keys,
      1,
    );
    assert.equal(
      (
        db.prepare('SELECT anime_id FROM imm_videos WHERE video_id = ?').get(videoId) as {
          anime_id: number;
        }
      ).anime_id,
      animeId,
    );
    assert.equal(
      (
        db
          .prepare('SELECT total_active_ms FROM imm_lifetime_anime WHERE anime_id = ?')
          .get(animeId) as { total_active_ms: number }
      ).total_active_ms,
      123,
    );
    db.prepare(
      "INSERT INTO imm_anime(normalized_title_key, canonical_title, media_kind) VALUES ('shared title', 'Channel', 'youtube')",
    ).run();
  } finally {
    db.close();
  }
});
