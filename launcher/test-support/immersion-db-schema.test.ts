import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Database as BunDatabase } from 'bun:sqlite';
import { ensureSchema } from '../../src/core/services/immersion-tracker/storage.js';
import { createImmersionDbFixture } from './immersion-db-fixture.js';

type SchemaRow = { type: string; name: string; tbl_name: string; sql: string | null };

const SYNC_SCHEMA_OBJECTS = [
  'imm_anime',
  'imm_videos',
  'imm_sessions',
  'imm_session_telemetry',
  'imm_session_events',
  'imm_daily_rollups',
  'imm_monthly_rollups',
  'imm_words',
  'imm_kanji',
  'imm_subtitle_lines',
  'imm_word_line_occurrences',
  'imm_kanji_line_occurrences',
  'imm_media_art',
  'imm_youtube_videos',
  'imm_cover_art_blobs',
  'imm_lifetime_global',
  'imm_lifetime_anime',
  'imm_lifetime_media',
  'imm_lifetime_applied_sessions',
  'imm_stats_excluded_words',
  'idx_anime_normalized_title',
  'idx_anime_anilist_id',
  'idx_videos_anime_id',
  'idx_sessions_video_started',
  'idx_sessions_status_started',
  'idx_sessions_started_at',
  'idx_sessions_ended_at',
  'idx_telemetry_session_sample',
  'idx_telemetry_sample_ms',
  'idx_events_session_ts',
  'idx_events_type_ts',
  'idx_rollups_day_video',
  'idx_rollups_month_video',
  'idx_words_headword_word_reading',
  'idx_words_frequency',
  'idx_kanji_kanji',
  'idx_kanji_frequency',
  'idx_subtitle_lines_session_line',
  'idx_subtitle_lines_video_line',
  'idx_subtitle_lines_anime_line',
  'idx_word_line_occurrences_word',
  'idx_kanji_line_occurrences_kanji',
  'idx_media_art_cover_blob_hash',
  'idx_media_art_anilist_id',
  'idx_media_art_cover_url',
  'idx_youtube_videos_channel_id',
  'idx_youtube_videos_youtube_video_id',
] as const;

function normalizeSql(sql: string | null): string {
  return (sql ?? '')
    .replace(/\bIF NOT EXISTS\b/gi, '')
    .replace(/\s+/g, ' ')
    .replace(/\s+([(),])/g, '$1')
    .replace(/,\s+/g, ', ')
    .trim();
}

function readSchema(dbPath: string): Map<string, string> {
  const db = new BunDatabase(dbPath, { readonly: true });
  try {
    const rows = db
      .query<SchemaRow>(
        `SELECT type, name, tbl_name, sql
         FROM sqlite_schema
         WHERE name IN (${SYNC_SCHEMA_OBJECTS.map(() => '?').join(',')})
         ORDER BY type, name`,
      )
      .all(...SYNC_SCHEMA_OBJECTS);
    return new Map(rows.map((row) => [row.name, normalizeSql(row.sql)]));
  } finally {
    db.close();
  }
}

test('fixture schema stays aligned with production sync-touched tables and indexes', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-fixture-schema-'));
  const fixturePath = path.join(dir, 'fixture.sqlite');
  const productionPath = path.join(dir, 'production.sqlite');
  const productionDb = new BunDatabase(productionPath, { create: true });
  try {
    createImmersionDbFixture(fixturePath);
    ensureSchema(productionDb as never);
    productionDb.close();

    const fixtureSchema = readSchema(fixturePath);
    const productionSchema = readSchema(productionPath);
    assert.deepEqual(fixtureSchema, productionSchema);
  } finally {
    try {
      productionDb.close();
    } catch {
      // already closed
    }
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
