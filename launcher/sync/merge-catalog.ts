import { Database } from 'bun:sqlite';
import { insertRow, tableExists, type SyncMergeSummary } from './sync-shared.js';

const ANIME_COPY_COLUMNS = [
  'normalized_title_key',
  'canonical_title',
  'anilist_id',
  'title_romaji',
  'title_english',
  'title_native',
  'episodes_total',
  'description',
  'metadata_json',
  'CREATED_DATE',
  'LAST_UPDATE_DATE',
] as const;

const VIDEO_COPY_COLUMNS = [
  'video_key',
  'canonical_title',
  'source_type',
  'source_path',
  'source_url',
  'parsed_basename',
  'parsed_title',
  'parsed_season',
  'parsed_episode',
  'parser_source',
  'parser_confidence',
  'parse_metadata_json',
  'watched',
  'duration_ms',
  'file_size_bytes',
  'codec_id',
  'container_id',
  'width_px',
  'height_px',
  'fps_x100',
  'bitrate_kbps',
  'audio_codec_id',
  'hash_sha256',
  'screenshot_path',
  'metadata_json',
  'CREATED_DATE',
  'LAST_UPDATE_DATE',
] as const;

const MEDIA_ART_COPY_COLUMNS = [
  'anilist_id',
  'cover_url',
  'cover_blob',
  'cover_blob_hash',
  'title_romaji',
  'title_english',
  'episodes_total',
  'fetched_at_ms',
  'CREATED_DATE',
  'LAST_UPDATE_DATE',
] as const;

const YOUTUBE_COPY_COLUMNS = [
  'youtube_video_id',
  'video_url',
  'video_title',
  'video_thumbnail_url',
  'channel_id',
  'channel_name',
  'channel_url',
  'channel_thumbnail_url',
  'uploader_id',
  'uploader_url',
  'description',
  'metadata_json',
  'fetched_at_ms',
  'CREATED_DATE',
  'LAST_UPDATE_DATE',
] as const;

const WORD_COPY_COLUMNS = [
  'headword',
  'word',
  'reading',
  'part_of_speech',
  'pos1',
  'pos2',
  'pos3',
  'first_seen',
  'last_seen',
  'frequency',
  'frequency_rank',
] as const;

type SqlRow = Record<string, unknown>;

function selectAll(db: Database, sql: string, params: unknown[] = []): SqlRow[] {
  return db.query<SqlRow>(sql).all(...params);
}

export function mergeAnime(
  local: Database,
  remote: Database,
  summary: SyncMergeSummary,
): Map<number, number> {
  const map = new Map<number, number>();
  const byAnilist = local.prepare<SqlRow>('SELECT anime_id FROM imm_anime WHERE anilist_id = ?');
  const byTitleKey = local.prepare<SqlRow>(
    'SELECT anime_id FROM imm_anime WHERE normalized_title_key = ?',
  );
  const fillMissing = local.prepare(
    `UPDATE imm_anime
     SET
       title_romaji = COALESCE(title_romaji, ?),
       title_english = COALESCE(title_english, ?),
       title_native = COALESCE(title_native, ?),
       episodes_total = COALESCE(episodes_total, ?),
       description = COALESCE(description, ?)
     WHERE anime_id = ?`,
  );

  for (const row of selectAll(
    remote,
    `SELECT anime_id, ${ANIME_COPY_COLUMNS.join(', ')} FROM imm_anime`,
  )) {
    const remoteId = Number(row.anime_id);
    const existing =
      (row.anilist_id !== null ? byAnilist.get(row.anilist_id) : undefined) ??
      byTitleKey.get(row.normalized_title_key);
    if (existing) {
      const localId = Number(existing.anime_id);
      map.set(remoteId, localId);
      fillMissing.run(
        row.title_romaji,
        row.title_english,
        row.title_native,
        row.episodes_total,
        row.description,
        localId,
      );
      continue;
    }
    // No local row matched by anilist_id (checked first in `existing` above)
    // or title key, so the remote anilist_id — if any — is free to insert as-is.
    const values = ANIME_COPY_COLUMNS.map((column) => row[column]);
    map.set(remoteId, insertRow(local, 'imm_anime', ANIME_COPY_COLUMNS, values));
    summary.animeAdded += 1;
  }
  return map;
}

export interface VideoMergeResult {
  videoIdMap: Map<number, number>;
  addedVideoIds: Set<number>;
}

export function mergeVideos(
  local: Database,
  remote: Database,
  animeIdMap: Map<number, number>,
  summary: SyncMergeSummary,
): VideoMergeResult {
  const videoIdMap = new Map<number, number>();
  const addedVideoIds = new Set<number>();
  const byKey = local.prepare<SqlRow>(
    'SELECT video_id, watched FROM imm_videos WHERE video_key = ?',
  );
  const setWatched = local.prepare('UPDATE imm_videos SET watched = 1 WHERE video_id = ?');

  for (const row of selectAll(
    remote,
    `SELECT video_id, anime_id, ${VIDEO_COPY_COLUMNS.join(', ')} FROM imm_videos`,
  )) {
    const remoteId = Number(row.video_id);
    const mappedAnimeId =
      row.anime_id === null ? null : (animeIdMap.get(Number(row.anime_id)) ?? null);
    const existing = byKey.get(row.video_key);
    if (existing) {
      const localId = Number(existing.video_id);
      videoIdMap.set(remoteId, localId);
      if (Number(row.watched) > 0 && Number(existing.watched) <= 0) {
        setWatched.run(localId);
      }
      continue;
    }
    const columns = ['anime_id', ...VIDEO_COPY_COLUMNS];
    const values = [mappedAnimeId, ...VIDEO_COPY_COLUMNS.map((column) => row[column])];
    const localId = insertRow(local, 'imm_videos', columns, values);
    videoIdMap.set(remoteId, localId);
    addedVideoIds.add(remoteId);
    summary.videosAdded += 1;
  }
  return { videoIdMap, addedVideoIds };
}

export function mergeMediaMetadata(
  local: Database,
  remote: Database,
  videoIdMap: Map<number, number>,
  addedVideoIds: Set<number>,
): void {
  if (videoIdMap.size === 0) return;
  const metadataVideoIds = new Set<number>([...addedVideoIds, ...videoIdMap.keys()]);

  const hasBlobStore =
    tableExists(local, 'imm_cover_art_blobs') && tableExists(remote, 'imm_cover_art_blobs');
  const copyBlob = hasBlobStore
    ? local.prepare(
        `INSERT INTO imm_cover_art_blobs (blob_hash, cover_blob, CREATED_DATE, LAST_UPDATE_DATE)
         VALUES (?, ?, ?, ?)
         ON CONFLICT(blob_hash) DO NOTHING`,
      )
    : null;
  const readBlob = hasBlobStore
    ? remote.prepare<SqlRow>('SELECT * FROM imm_cover_art_blobs WHERE blob_hash = ?')
    : null;

  if (tableExists(remote, 'imm_media_art') && tableExists(local, 'imm_media_art')) {
    const localArtExists = local.prepare<SqlRow>(
      'SELECT 1 FROM imm_media_art WHERE video_id = ? LIMIT 1',
    );
    for (const remoteVideoId of metadataVideoIds) {
      const localVideoId = videoIdMap.get(remoteVideoId)!;
      if (localArtExists.get(localVideoId)) continue;
      const row = remote
        .query<SqlRow>(
          `SELECT ${MEDIA_ART_COPY_COLUMNS.join(', ')} FROM imm_media_art WHERE video_id = ?`,
        )
        .get(remoteVideoId);
      if (!row) continue;
      if (row.cover_blob_hash && copyBlob && readBlob) {
        const blob = readBlob.get(row.cover_blob_hash);
        if (blob) {
          copyBlob.run(blob.blob_hash, blob.cover_blob, blob.CREATED_DATE, blob.LAST_UPDATE_DATE);
        }
      }
      insertRow(
        local,
        'imm_media_art',
        ['video_id', ...MEDIA_ART_COPY_COLUMNS],
        [localVideoId, ...MEDIA_ART_COPY_COLUMNS.map((column) => row[column])],
      );
    }
  }

  if (tableExists(remote, 'imm_youtube_videos') && tableExists(local, 'imm_youtube_videos')) {
    const localYoutubeExists = local.prepare<SqlRow>(
      'SELECT 1 FROM imm_youtube_videos WHERE video_id = ? LIMIT 1',
    );
    for (const remoteVideoId of metadataVideoIds) {
      const localVideoId = videoIdMap.get(remoteVideoId)!;
      if (localYoutubeExists.get(localVideoId)) continue;
      const row = remote
        .query<SqlRow>(
          `SELECT ${YOUTUBE_COPY_COLUMNS.join(', ')} FROM imm_youtube_videos WHERE video_id = ?`,
        )
        .get(remoteVideoId);
      if (!row) continue;
      insertRow(
        local,
        'imm_youtube_videos',
        ['video_id', ...YOUTUBE_COPY_COLUMNS],
        [localVideoId, ...YOUTUBE_COPY_COLUMNS.map((column) => row[column])],
      );
    }
  }
}

export function mergeExcludedWords(
  local: Database,
  remote: Database,
  summary: SyncMergeSummary,
): void {
  if (
    !tableExists(remote, 'imm_stats_excluded_words') ||
    !tableExists(local, 'imm_stats_excluded_words')
  ) {
    return;
  }
  const insert = local.prepare(
    `INSERT INTO imm_stats_excluded_words (headword, word, reading, CREATED_DATE, LAST_UPDATE_DATE)
     VALUES (?, ?, ?, ?, ?)
     ON CONFLICT(headword, word, reading) DO NOTHING`,
  );
  for (const row of selectAll(
    remote,
    'SELECT headword, word, reading, CREATED_DATE, LAST_UPDATE_DATE FROM imm_stats_excluded_words',
  )) {
    const result = insert.run(
      row.headword,
      row.word,
      row.reading,
      row.CREATED_DATE,
      row.LAST_UPDATE_DATE,
    );
    summary.excludedWordsAdded += result.changes;
  }
}

/**
 * Lazily maps remote imm_words / imm_kanji ids onto local rows by natural key
 * ((headword, word, reading) / kanji). New rows are copied with the remote's
 * accumulated frequency; rows that already exist locally get their frequency
 * incremented later with only the occurrence counts this merge adds (the
 * remote total would double-count lines merged in earlier syncs).
 */
export class LexiconResolver {
  private readonly wordMap = new Map<number, { localId: number; isNew: boolean }>();
  private readonly kanjiMap = new Map<number, { localId: number; isNew: boolean }>();
  readonly wordFrequencyDeltas = new Map<number, number>();
  readonly kanjiFrequencyDeltas = new Map<number, number>();

  constructor(
    private readonly local: Database,
    private readonly remote: Database,
    private readonly summary: SyncMergeSummary,
  ) {}

  resolveWord(remoteWordId: number): number {
    const cached = this.wordMap.get(remoteWordId);
    if (cached) return cached.localId;

    const row = this.remote
      .query<SqlRow>(`SELECT ${WORD_COPY_COLUMNS.join(', ')} FROM imm_words WHERE id = ?`)
      .get(remoteWordId);
    if (!row) throw new Error(`Snapshot references missing imm_words row ${remoteWordId}`);

    const existing = this.local
      .query<SqlRow>('SELECT id FROM imm_words WHERE headword IS ? AND word IS ? AND reading IS ?')
      .get(row.headword, row.word, row.reading);
    let entry: { localId: number; isNew: boolean };
    if (existing) {
      entry = { localId: Number(existing.id), isNew: false };
      this.local
        .prepare(
          `UPDATE imm_words
           SET first_seen = MIN(COALESCE(first_seen, ?), COALESCE(?, first_seen)),
               last_seen = MAX(COALESCE(last_seen, ?), COALESCE(?, last_seen))
           WHERE id = ?`,
        )
        .run(row.first_seen, row.first_seen, row.last_seen, row.last_seen, entry.localId);
    } else {
      const localId = insertRow(
        this.local,
        'imm_words',
        WORD_COPY_COLUMNS,
        WORD_COPY_COLUMNS.map((column) => row[column]),
      );
      entry = { localId, isNew: true };
      this.summary.wordsAdded += 1;
    }
    this.wordMap.set(remoteWordId, entry);
    return entry.localId;
  }

  resolveKanji(remoteKanjiId: number): number {
    const cached = this.kanjiMap.get(remoteKanjiId);
    if (cached) return cached.localId;

    const row = this.remote
      .query<SqlRow>('SELECT kanji, first_seen, last_seen, frequency FROM imm_kanji WHERE id = ?')
      .get(remoteKanjiId);
    if (!row) throw new Error(`Snapshot references missing imm_kanji row ${remoteKanjiId}`);

    const existing = this.local
      .query<SqlRow>('SELECT id FROM imm_kanji WHERE kanji IS ?')
      .get(row.kanji);
    let entry: { localId: number; isNew: boolean };
    if (existing) {
      entry = { localId: Number(existing.id), isNew: false };
      this.local
        .prepare(
          `UPDATE imm_kanji
           SET first_seen = MIN(COALESCE(first_seen, ?), COALESCE(?, first_seen)),
               last_seen = MAX(COALESCE(last_seen, ?), COALESCE(?, last_seen))
           WHERE id = ?`,
        )
        .run(row.first_seen, row.first_seen, row.last_seen, row.last_seen, entry.localId);
    } else {
      const localId = insertRow(
        this.local,
        'imm_kanji',
        ['kanji', 'first_seen', 'last_seen', 'frequency'],
        [row.kanji, row.first_seen, row.last_seen, row.frequency],
      );
      entry = { localId, isNew: true };
      this.summary.kanjiAdded += 1;
    }
    this.kanjiMap.set(remoteKanjiId, entry);
    return entry.localId;
  }

  addWordOccurrences(remoteWordId: number, count: number): void {
    const entry = this.wordMap.get(remoteWordId);
    if (!entry || entry.isNew) return;
    this.wordFrequencyDeltas.set(
      entry.localId,
      (this.wordFrequencyDeltas.get(entry.localId) ?? 0) + count,
    );
  }

  addKanjiOccurrences(remoteKanjiId: number, count: number): void {
    const entry = this.kanjiMap.get(remoteKanjiId);
    if (!entry || entry.isNew) return;
    this.kanjiFrequencyDeltas.set(
      entry.localId,
      (this.kanjiFrequencyDeltas.get(entry.localId) ?? 0) + count,
    );
  }

  applyFrequencyDeltas(): void {
    const updateWord = this.local.prepare(
      'UPDATE imm_words SET frequency = COALESCE(frequency, 0) + ? WHERE id = ?',
    );
    for (const [localId, delta] of this.wordFrequencyDeltas) {
      updateWord.run(delta, localId);
    }
    const updateKanji = this.local.prepare(
      'UPDATE imm_kanji SET frequency = COALESCE(frequency, 0) + ? WHERE id = ?',
    );
    for (const [localId, delta] of this.kanjiFrequencyDeltas) {
      updateKanji.run(delta, localId);
    }
  }
}
