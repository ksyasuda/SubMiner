import { selectAll, selectOne, type SqlRow, type SyncDb } from './libsql-driver';
import { insertRow, tableExists, type SyncMergeSummary } from './shared';
import { sameTitleNamespaceSql } from '../../../shared/media-kind';

const ANIME_COPY_COLUMNS = [
  'media_kind',
  'normalized_title_key',
  'canonical_title',
  'anilist_id',
  'title_romaji',
  'title_english',
  'title_native',
  'episodes_total',
  'description',
  'tmdb_id',
  'tmdb_type',
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
  'anime_assignment_locked',
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
  'vocabulary_visible',
] as const;

export function mergeAnime(
  local: SyncDb,
  remote: SyncDb,
  summary: SyncMergeSummary,
): Map<number, number> {
  const map = new Map<number, number>();
  const byAnilist = local.query(
    "SELECT anime_id FROM imm_anime WHERE anilist_id = ? AND media_kind = 'anime'",
  );
  const byTmdb = local.query(
    'SELECT anime_id FROM imm_anime WHERE tmdb_id = ? AND tmdb_type = ? ORDER BY anime_id LIMIT 1',
  );
  // Anime and live-action rows share a title namespace; YouTube channels are
  // looked up on their own, so a same-named anime and channel stay separate.
  const byTitleKey = local.query(
    `SELECT anime_id, anilist_id, tmdb_id, tmdb_type FROM imm_anime
     WHERE normalized_title_key = ? AND ${sameTitleNamespaceSql()}`,
  );
  // A pre-classification channel can be repaired, but a genuine anime sharing
  // its title must remain a separate entry.
  const legacyChannel = local.query(`SELECT anime_id FROM imm_anime
    WHERE normalized_title_key = ? AND media_kind = 'anime' AND (
      normalized_title_key LIKE 'youtube channel %'
      OR CASE WHEN json_valid(metadata_json)
        THEN json_extract(metadata_json, '$.source') = 'youtube-channel' ELSE 0 END
    )`);
  const releaseChannelAnilistId = local.query(
    "UPDATE imm_anime SET anilist_id = NULL WHERE media_kind = 'youtube' AND anilist_id = ?",
  );
  // A TMDB link only fills in when the local row is unlinked: a row already
  // pinned to AniList stays anime, and vice versa, so the two link kinds never
  // coexist on one entry. A channel match always becomes a channel.
  const fillMissing = local.query(
    `UPDATE imm_anime
     SET
       anilist_id = CASE WHEN ? = 'youtube' THEN NULL ELSE anilist_id END,
       title_romaji = COALESCE(title_romaji, ?),
       title_english = COALESCE(title_english, ?),
       title_native = COALESCE(title_native, ?),
       episodes_total = COALESCE(episodes_total, ?),
       description = COALESCE(description, ?),
       tmdb_id = CASE WHEN anilist_id IS NULL THEN COALESCE(tmdb_id, ?) ELSE tmdb_id END,
       tmdb_type = CASE WHEN anilist_id IS NULL AND tmdb_id IS NULL THEN ? ELSE tmdb_type END,
       media_kind = CASE
         WHEN ? = 'youtube' THEN 'youtube'
         WHEN anilist_id IS NULL AND tmdb_id IS NULL THEN ?
         ELSE media_kind
       END
     WHERE anime_id = ?`,
  );

  for (const row of selectAll(
    remote,
    `SELECT anime_id, ${ANIME_COPY_COLUMNS.join(', ')} FROM imm_anime`,
  )) {
    const remoteId = Number(row.anime_id);
    if (row.media_kind === 'anime' && row.anilist_id !== null) {
      // AniList identifiers belong to anime, including when an older peer
      // incorrectly attached one to a channel.
      releaseChannelAnilistId.run(row.anilist_id);
    }
    const titleMatch = byTitleKey.get(row.normalized_title_key, row.media_kind) as
      | SqlRow
      | undefined;
    const compatibleTitleMatch =
      titleMatch &&
      ((titleMatch.anilist_id === null && titleMatch.tmdb_id === null) ||
        (row.anilist_id === null && row.tmdb_id === null) ||
        (titleMatch.tmdb_id === null &&
          row.tmdb_id === null &&
          titleMatch.anilist_id === row.anilist_id) ||
        (titleMatch.anilist_id === null &&
          row.anilist_id === null &&
          titleMatch.tmdb_id === row.tmdb_id &&
          titleMatch.tmdb_type === row.tmdb_type));
    const existing = ((row.media_kind === 'anime' && row.anilist_id !== null
      ? byAnilist.get(row.anilist_id)
      : undefined) ??
      (row.tmdb_id !== null ? byTmdb.get(row.tmdb_id, row.tmdb_type) : undefined) ??
      (compatibleTitleMatch ? titleMatch : undefined) ??
      (row.media_kind === 'youtube' ? legacyChannel.get(row.normalized_title_key) : undefined)) as
      | SqlRow
      | undefined;
    if (existing) {
      const localId = Number(existing.anime_id);
      map.set(remoteId, localId);
      fillMissing.run(
        row.media_kind,
        row.title_romaji,
        row.title_english,
        row.title_native,
        row.episodes_total,
        row.description,
        row.tmdb_id,
        row.tmdb_type,
        row.media_kind,
        row.media_kind,
        localId,
      );
      continue;
    }
    // Conflicting providers can share a title, but the stored title key is
    // unique within its namespace.
    let titleKey = row.normalized_title_key;
    for (let suffix = 1; byTitleKey.get(titleKey, row.media_kind); suffix += 1) {
      titleKey = `${row.normalized_title_key}:sync:${suffix}`;
    }
    const values = ANIME_COPY_COLUMNS.map((column) => {
      if (column === 'normalized_title_key') return titleKey;
      // No local row matched by anilist_id (checked first in `existing` above)
      // or title key, so the remote anilist_id is free to insert as-is, except
      // that channels never carry one.
      if (column === 'anilist_id' && row.media_kind !== 'anime') return null;
      return row[column];
    });
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
  local: SyncDb,
  remote: SyncDb,
  animeIdMap: Map<number, number>,
  summary: SyncMergeSummary,
): VideoMergeResult {
  const videoIdMap = new Map<number, number>();
  const addedVideoIds = new Set<number>();
  const byKey = local.query('SELECT video_id, watched FROM imm_videos WHERE video_key = ?');
  const setWatched = local.query('UPDATE imm_videos SET watched = 1 WHERE video_id = ?');

  for (const row of selectAll(
    remote,
    `SELECT video_id, anime_id, ${VIDEO_COPY_COLUMNS.join(', ')} FROM imm_videos`,
  )) {
    const remoteId = Number(row.video_id);
    const mappedAnimeId =
      row.anime_id === null ? null : (animeIdMap.get(Number(row.anime_id)) ?? null);
    const existing = byKey.get(row.video_key) as SqlRow | undefined;
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
  local: SyncDb,
  remote: SyncDb,
  videoIdMap: Map<number, number>,
  addedVideoIds: Set<number>,
): void {
  if (videoIdMap.size === 0) return;
  const metadataVideoIds = new Set<number>([...addedVideoIds, ...videoIdMap.keys()]);

  const hasBlobStore =
    tableExists(local, 'imm_cover_art_blobs') && tableExists(remote, 'imm_cover_art_blobs');
  const copyBlob = hasBlobStore
    ? local.query(
        `INSERT INTO imm_cover_art_blobs (blob_hash, cover_blob, CREATED_DATE, LAST_UPDATE_DATE)
         VALUES (?, ?, ?, ?)
         ON CONFLICT(blob_hash) DO NOTHING`,
      )
    : null;
  const readBlob = hasBlobStore
    ? remote.query('SELECT * FROM imm_cover_art_blobs WHERE blob_hash = ?')
    : null;

  if (tableExists(remote, 'imm_media_art') && tableExists(local, 'imm_media_art')) {
    const localArtExists = local.query('SELECT 1 FROM imm_media_art WHERE video_id = ? LIMIT 1');
    for (const remoteVideoId of metadataVideoIds) {
      const localVideoId = videoIdMap.get(remoteVideoId)!;
      if (localArtExists.get(localVideoId)) continue;
      const row = selectOne(
        remote,
        `SELECT ${MEDIA_ART_COPY_COLUMNS.join(', ')} FROM imm_media_art WHERE video_id = ?`,
        [remoteVideoId],
      );
      if (!row) continue;
      if (row.cover_blob_hash && copyBlob && readBlob) {
        const blob = readBlob.get(row.cover_blob_hash) as SqlRow | undefined;
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
    const localYoutubeExists = local.query(
      'SELECT 1 FROM imm_youtube_videos WHERE video_id = ? LIMIT 1',
    );
    for (const remoteVideoId of metadataVideoIds) {
      const localVideoId = videoIdMap.get(remoteVideoId)!;
      if (localYoutubeExists.get(localVideoId)) continue;
      const row = selectOne(
        remote,
        `SELECT ${YOUTUBE_COPY_COLUMNS.join(', ')} FROM imm_youtube_videos WHERE video_id = ?`,
        [remoteVideoId],
      );
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

export function mergeExcludedWords(local: SyncDb, remote: SyncDb, summary: SyncMergeSummary): void {
  if (
    !tableExists(remote, 'imm_stats_excluded_words') ||
    !tableExists(local, 'imm_stats_excluded_words')
  ) {
    return;
  }
  const insert = local.query(
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
 * accumulated frequency minus any counts owed to skipped ACTIVE sessions
 * (those merge later and re-add their counts); rows that already exist locally
 * get their frequency incremented later with only the occurrence counts this
 * merge adds (the remote total would double-count lines merged in earlier
 * syncs).
 */
export class LexiconResolver {
  private readonly wordMap = new Map<number, { localId: number; isNew: boolean }>();
  private readonly kanjiMap = new Map<number, { localId: number; isNew: boolean }>();
  readonly wordFrequencyDeltas = new Map<number, number>();
  readonly kanjiFrequencyDeltas = new Map<number, number>();

  constructor(
    private readonly local: SyncDb,
    private readonly remote: SyncDb,
    private readonly summary: SyncMergeSummary,
  ) {}

  /**
   * Occurrence counts the remote's live tracker already baked into `frequency`
   * but that belong to skipped ACTIVE sessions. Those lines are not copied this
   * merge; they are re-added via addWordOccurrences/addKanjiOccurrences when
   * the session finalizes and syncs, so a newly adopted row must not carry them
   * or that slice would be counted twice.
   */
  private pendingActiveSessionOccurrences(
    occurrenceTable: 'imm_word_line_occurrences' | 'imm_kanji_line_occurrences',
    idColumn: 'word_id' | 'kanji_id',
    remoteId: number,
  ): number {
    const row = selectOne(
      this.remote,
      `SELECT COALESCE(SUM(o.occurrence_count), 0) AS pending
       FROM ${occurrenceTable} o
       JOIN imm_subtitle_lines l ON l.line_id = o.line_id
       JOIN imm_sessions s ON s.session_id = l.session_id
       WHERE o.${idColumn} = ? AND s.ended_at_ms IS NULL`,
      [remoteId],
    );
    return Number(row?.pending ?? 0);
  }

  private adoptedFrequency(frequency: unknown, pending: number): unknown {
    if (pending <= 0 || typeof frequency !== 'number') return frequency;
    return Math.max(0, frequency - pending);
  }

  resolveWord(remoteWordId: number): number {
    const cached = this.wordMap.get(remoteWordId);
    if (cached) return cached.localId;

    const row = selectOne(
      this.remote,
      `SELECT ${WORD_COPY_COLUMNS.join(', ')} FROM imm_words WHERE id = ?`,
      [remoteWordId],
    );
    if (!row) throw new Error(`Snapshot references missing imm_words row ${remoteWordId}`);

    const existing = selectOne(
      this.local,
      'SELECT id FROM imm_words WHERE headword IS ? AND word IS ? AND reading IS ?',
      [row.headword, row.word, row.reading],
    );
    let entry: { localId: number; isNew: boolean };
    if (existing) {
      entry = { localId: Number(existing.id), isNew: false };
      this.local
        .query(
          `UPDATE imm_words
           SET first_seen = MIN(COALESCE(first_seen, ?), COALESCE(?, first_seen)),
               last_seen = MAX(COALESCE(last_seen, ?), COALESCE(?, last_seen))
           WHERE id = ?`,
        )
        .run(row.first_seen, row.first_seen, row.last_seen, row.last_seen, entry.localId);
    } else {
      const pending = this.pendingActiveSessionOccurrences(
        'imm_word_line_occurrences',
        'word_id',
        remoteWordId,
      );
      const localId = insertRow(
        this.local,
        'imm_words',
        WORD_COPY_COLUMNS,
        WORD_COPY_COLUMNS.map((column) =>
          column === 'frequency' ? this.adoptedFrequency(row[column], pending) : row[column],
        ),
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

    const row = selectOne(
      this.remote,
      'SELECT kanji, first_seen, last_seen, frequency FROM imm_kanji WHERE id = ?',
      [remoteKanjiId],
    );
    if (!row) throw new Error(`Snapshot references missing imm_kanji row ${remoteKanjiId}`);

    const existing = selectOne(this.local, 'SELECT id FROM imm_kanji WHERE kanji IS ?', [
      row.kanji,
    ]);
    let entry: { localId: number; isNew: boolean };
    if (existing) {
      entry = { localId: Number(existing.id), isNew: false };
      this.local
        .query(
          `UPDATE imm_kanji
           SET first_seen = MIN(COALESCE(first_seen, ?), COALESCE(?, first_seen)),
               last_seen = MAX(COALESCE(last_seen, ?), COALESCE(?, last_seen))
           WHERE id = ?`,
        )
        .run(row.first_seen, row.first_seen, row.last_seen, row.last_seen, entry.localId);
    } else {
      const pending = this.pendingActiveSessionOccurrences(
        'imm_kanji_line_occurrences',
        'kanji_id',
        remoteKanjiId,
      );
      const localId = insertRow(
        this.local,
        'imm_kanji',
        ['kanji', 'first_seen', 'last_seen', 'frequency'],
        [row.kanji, row.first_seen, row.last_seen, this.adoptedFrequency(row.frequency, pending)],
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
    const updateWord = this.local.query(
      'UPDATE imm_words SET frequency = COALESCE(frequency, 0) + ? WHERE id = ?',
    );
    for (const [localId, delta] of this.wordFrequencyDeltas) {
      updateWord.run(delta, localId);
    }
    const updateKanji = this.local.query(
      'UPDATE imm_kanji SET frequency = COALESCE(frequency, 0) + ? WHERE id = ?',
    );
    for (const [localId, delta] of this.kanjiFrequencyDeltas) {
      updateKanji.run(delta, localId);
    }
  }
}
