// Schema-version-18 shape of the tables the sync merge touches (plus the
// app's indexes), mirroring ensureSchema / ensureLifetimeSummaryTables /
// ensureStatsExcludedWordsTable in src/core/services/immersion-tracker/storage.ts.
export const IMMERSION_DB_FIXTURE_DDL = `
  CREATE TABLE imm_schema_version (
    schema_version INTEGER PRIMARY KEY,
    applied_at_ms TEXT NOT NULL
  );
  CREATE TABLE imm_rollup_state(
    state_key TEXT PRIMARY KEY,
    state_value TEXT NOT NULL
  );
  CREATE TABLE imm_anime(
    anime_id INTEGER PRIMARY KEY AUTOINCREMENT,
    normalized_title_key TEXT NOT NULL UNIQUE,
    canonical_title TEXT NOT NULL,
    anilist_id INTEGER UNIQUE,
    title_romaji TEXT,
    title_english TEXT,
    title_native TEXT,
    episodes_total INTEGER,
    description TEXT,
    metadata_json TEXT,
    CREATED_DATE TEXT,
    LAST_UPDATE_DATE TEXT
  );
  CREATE TABLE imm_videos(
    video_id INTEGER PRIMARY KEY AUTOINCREMENT,
    video_key TEXT NOT NULL UNIQUE,
    anime_id INTEGER,
    canonical_title TEXT NOT NULL,
    source_type INTEGER NOT NULL,
    source_path TEXT,
    source_url TEXT,
    parsed_basename TEXT,
    parsed_title TEXT,
    parsed_season INTEGER,
    parsed_episode INTEGER,
    parser_source TEXT,
    parser_confidence REAL,
    parse_metadata_json TEXT,
    watched INTEGER NOT NULL DEFAULT 0,
    duration_ms INTEGER NOT NULL CHECK(duration_ms>=0),
    file_size_bytes INTEGER CHECK(file_size_bytes>=0),
    codec_id INTEGER, container_id INTEGER,
    width_px INTEGER, height_px INTEGER, fps_x100 INTEGER,
    bitrate_kbps INTEGER, audio_codec_id INTEGER,
    hash_sha256 TEXT, screenshot_path TEXT,
    metadata_json TEXT,
    CREATED_DATE TEXT,
    LAST_UPDATE_DATE TEXT,
    FOREIGN KEY(anime_id) REFERENCES imm_anime(anime_id) ON DELETE SET NULL
  );
  CREATE TABLE imm_sessions(
    session_id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_uuid TEXT NOT NULL UNIQUE,
    video_id INTEGER NOT NULL,
    started_at_ms TEXT NOT NULL, ended_at_ms TEXT,
    status INTEGER NOT NULL,
    locale_id INTEGER, target_lang_id INTEGER,
    difficulty_tier INTEGER, subtitle_mode INTEGER,
    ended_media_ms INTEGER,
    total_watched_ms INTEGER NOT NULL DEFAULT 0,
    active_watched_ms INTEGER NOT NULL DEFAULT 0,
    lines_seen INTEGER NOT NULL DEFAULT 0,
    tokens_seen INTEGER NOT NULL DEFAULT 0,
    cards_mined INTEGER NOT NULL DEFAULT 0,
    lookup_count INTEGER NOT NULL DEFAULT 0,
    lookup_hits INTEGER NOT NULL DEFAULT 0,
    yomitan_lookup_count INTEGER NOT NULL DEFAULT 0,
    pause_count INTEGER NOT NULL DEFAULT 0,
    pause_ms INTEGER NOT NULL DEFAULT 0,
    seek_forward_count INTEGER NOT NULL DEFAULT 0,
    seek_backward_count INTEGER NOT NULL DEFAULT 0,
    media_buffer_events INTEGER NOT NULL DEFAULT 0,
    CREATED_DATE TEXT,
    LAST_UPDATE_DATE TEXT,
    FOREIGN KEY(video_id) REFERENCES imm_videos(video_id)
  );
  CREATE TABLE imm_session_telemetry(
    telemetry_id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id INTEGER NOT NULL,
    sample_ms TEXT NOT NULL,
    total_watched_ms INTEGER NOT NULL DEFAULT 0,
    active_watched_ms INTEGER NOT NULL DEFAULT 0,
    lines_seen INTEGER NOT NULL DEFAULT 0,
    tokens_seen INTEGER NOT NULL DEFAULT 0,
    cards_mined INTEGER NOT NULL DEFAULT 0,
    lookup_count INTEGER NOT NULL DEFAULT 0,
    lookup_hits INTEGER NOT NULL DEFAULT 0,
    yomitan_lookup_count INTEGER NOT NULL DEFAULT 0,
    pause_count INTEGER NOT NULL DEFAULT 0,
    pause_ms INTEGER NOT NULL DEFAULT 0,
    seek_forward_count INTEGER NOT NULL DEFAULT 0,
    seek_backward_count INTEGER NOT NULL DEFAULT 0,
    media_buffer_events INTEGER NOT NULL DEFAULT 0,
    CREATED_DATE TEXT,
    LAST_UPDATE_DATE TEXT,
    FOREIGN KEY(session_id) REFERENCES imm_sessions(session_id) ON DELETE CASCADE
  );
  CREATE TABLE imm_session_events(
    event_id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id INTEGER NOT NULL,
    ts_ms TEXT NOT NULL,
    event_type INTEGER NOT NULL,
    line_index INTEGER,
    segment_start_ms INTEGER,
    segment_end_ms INTEGER,
    tokens_delta INTEGER NOT NULL DEFAULT 0,
    cards_delta INTEGER NOT NULL DEFAULT 0,
    payload_json TEXT,
    CREATED_DATE TEXT,
    LAST_UPDATE_DATE TEXT,
    FOREIGN KEY(session_id) REFERENCES imm_sessions(session_id) ON DELETE CASCADE
  );
  CREATE TABLE imm_daily_rollups(
    rollup_day INTEGER NOT NULL,
    video_id INTEGER,
    total_sessions INTEGER NOT NULL DEFAULT 0,
    total_active_min REAL NOT NULL DEFAULT 0,
    total_lines_seen INTEGER NOT NULL DEFAULT 0,
    total_tokens_seen INTEGER NOT NULL DEFAULT 0,
    total_cards INTEGER NOT NULL DEFAULT 0,
    cards_per_hour REAL,
    tokens_per_min REAL,
    lookup_hit_rate REAL,
    CREATED_DATE TEXT,
    LAST_UPDATE_DATE TEXT,
    PRIMARY KEY (rollup_day, video_id)
  );
  CREATE TABLE imm_monthly_rollups(
    rollup_month INTEGER NOT NULL,
    video_id INTEGER,
    total_sessions INTEGER NOT NULL DEFAULT 0,
    total_active_min REAL NOT NULL DEFAULT 0,
    total_lines_seen INTEGER NOT NULL DEFAULT 0,
    total_tokens_seen INTEGER NOT NULL DEFAULT 0,
    total_cards INTEGER NOT NULL DEFAULT 0,
    CREATED_DATE TEXT,
    LAST_UPDATE_DATE TEXT,
    PRIMARY KEY (rollup_month, video_id)
  );
  CREATE TABLE imm_words(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    headword TEXT,
    word TEXT,
    reading TEXT,
    part_of_speech TEXT,
    pos1 TEXT,
    pos2 TEXT,
    pos3 TEXT,
    first_seen REAL,
    last_seen REAL,
    frequency INTEGER,
    frequency_rank INTEGER,
    UNIQUE(headword, word, reading)
  );
  CREATE TABLE imm_kanji(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    kanji TEXT,
    first_seen REAL,
    last_seen REAL,
    frequency INTEGER,
    UNIQUE(kanji)
  );
  CREATE TABLE imm_subtitle_lines(
    line_id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id INTEGER NOT NULL,
    event_id INTEGER,
    video_id INTEGER NOT NULL,
    anime_id INTEGER,
    line_index INTEGER NOT NULL,
    segment_start_ms INTEGER,
    segment_end_ms INTEGER,
    text TEXT NOT NULL,
    secondary_text TEXT,
    CREATED_DATE INTEGER,
    LAST_UPDATE_DATE INTEGER,
    FOREIGN KEY(session_id) REFERENCES imm_sessions(session_id) ON DELETE CASCADE,
    FOREIGN KEY(event_id) REFERENCES imm_session_events(event_id) ON DELETE SET NULL,
    FOREIGN KEY(video_id) REFERENCES imm_videos(video_id) ON DELETE CASCADE,
    FOREIGN KEY(anime_id) REFERENCES imm_anime(anime_id) ON DELETE SET NULL
  );
  CREATE TABLE imm_word_line_occurrences(
    line_id INTEGER NOT NULL,
    word_id INTEGER NOT NULL,
    occurrence_count INTEGER NOT NULL,
    PRIMARY KEY(line_id, word_id),
    FOREIGN KEY(line_id) REFERENCES imm_subtitle_lines(line_id) ON DELETE CASCADE,
    FOREIGN KEY(word_id) REFERENCES imm_words(id) ON DELETE CASCADE
  );
  CREATE TABLE imm_kanji_line_occurrences(
    line_id INTEGER NOT NULL,
    kanji_id INTEGER NOT NULL,
    occurrence_count INTEGER NOT NULL,
    PRIMARY KEY(line_id, kanji_id),
    FOREIGN KEY(line_id) REFERENCES imm_subtitle_lines(line_id) ON DELETE CASCADE,
    FOREIGN KEY(kanji_id) REFERENCES imm_kanji(id) ON DELETE CASCADE
  );
  CREATE TABLE imm_media_art(
    video_id INTEGER PRIMARY KEY,
    anilist_id INTEGER,
    cover_url TEXT,
    cover_blob BLOB,
    cover_blob_hash TEXT,
    title_romaji TEXT,
    title_english TEXT,
    episodes_total INTEGER,
    fetched_at_ms TEXT NOT NULL,
    CREATED_DATE TEXT,
    LAST_UPDATE_DATE TEXT,
    FOREIGN KEY(video_id) REFERENCES imm_videos(video_id) ON DELETE CASCADE
  );
  CREATE TABLE imm_youtube_videos(
    video_id INTEGER PRIMARY KEY,
    youtube_video_id TEXT NOT NULL,
    video_url TEXT NOT NULL,
    video_title TEXT,
    video_thumbnail_url TEXT,
    channel_id TEXT,
    channel_name TEXT,
    channel_url TEXT,
    channel_thumbnail_url TEXT,
    uploader_id TEXT,
    uploader_url TEXT,
    description TEXT,
    metadata_json TEXT,
    fetched_at_ms TEXT NOT NULL,
    CREATED_DATE TEXT,
    LAST_UPDATE_DATE TEXT,
    FOREIGN KEY(video_id) REFERENCES imm_videos(video_id) ON DELETE CASCADE
  );
  CREATE TABLE imm_cover_art_blobs(
    blob_hash TEXT PRIMARY KEY,
    cover_blob BLOB NOT NULL,
    CREATED_DATE TEXT,
    LAST_UPDATE_DATE TEXT
  );
  CREATE TABLE imm_lifetime_global(
    global_id INTEGER PRIMARY KEY CHECK(global_id = 1),
    total_sessions INTEGER NOT NULL DEFAULT 0,
    total_active_ms INTEGER NOT NULL DEFAULT 0,
    total_cards INTEGER NOT NULL DEFAULT 0,
    active_days INTEGER NOT NULL DEFAULT 0,
    episodes_started INTEGER NOT NULL DEFAULT 0,
    episodes_completed INTEGER NOT NULL DEFAULT 0,
    anime_completed INTEGER NOT NULL DEFAULT 0,
    last_rebuilt_ms TEXT,
    CREATED_DATE TEXT,
    LAST_UPDATE_DATE TEXT
  );
  CREATE TABLE imm_lifetime_anime(
    anime_id INTEGER PRIMARY KEY,
    total_sessions INTEGER NOT NULL DEFAULT 0,
    total_active_ms INTEGER NOT NULL DEFAULT 0,
    total_cards INTEGER NOT NULL DEFAULT 0,
    total_lines_seen INTEGER NOT NULL DEFAULT 0,
    total_tokens_seen INTEGER NOT NULL DEFAULT 0,
    episodes_started INTEGER NOT NULL DEFAULT 0,
    episodes_completed INTEGER NOT NULL DEFAULT 0,
    first_watched_ms TEXT,
    last_watched_ms TEXT,
    CREATED_DATE TEXT,
    LAST_UPDATE_DATE TEXT,
    FOREIGN KEY(anime_id) REFERENCES imm_anime(anime_id) ON DELETE CASCADE
  );
  CREATE TABLE imm_lifetime_media(
    video_id INTEGER PRIMARY KEY,
    total_sessions INTEGER NOT NULL DEFAULT 0,
    total_active_ms INTEGER NOT NULL DEFAULT 0,
    total_cards INTEGER NOT NULL DEFAULT 0,
    total_lines_seen INTEGER NOT NULL DEFAULT 0,
    total_tokens_seen INTEGER NOT NULL DEFAULT 0,
    completed INTEGER NOT NULL DEFAULT 0,
    first_watched_ms TEXT,
    last_watched_ms TEXT,
    CREATED_DATE TEXT,
    LAST_UPDATE_DATE TEXT,
    FOREIGN KEY(video_id) REFERENCES imm_videos(video_id) ON DELETE CASCADE
  );
  CREATE TABLE imm_lifetime_applied_sessions(
    session_id INTEGER PRIMARY KEY,
    applied_at_ms TEXT NOT NULL,
    CREATED_DATE TEXT,
    LAST_UPDATE_DATE TEXT,
    FOREIGN KEY(session_id) REFERENCES imm_sessions(session_id) ON DELETE CASCADE
  );
  CREATE TABLE imm_stats_excluded_words(
    headword TEXT NOT NULL,
    word TEXT NOT NULL,
    reading TEXT NOT NULL,
    CREATED_DATE TEXT,
    LAST_UPDATE_DATE TEXT,
    PRIMARY KEY(headword, word, reading)
  );
  CREATE INDEX idx_anime_normalized_title ON imm_anime(normalized_title_key);
  CREATE INDEX idx_anime_anilist_id ON imm_anime(anilist_id);
  CREATE INDEX idx_videos_anime_id ON imm_videos(anime_id);
  CREATE INDEX idx_sessions_video_started ON imm_sessions(video_id, started_at_ms DESC);
  CREATE INDEX idx_sessions_status_started ON imm_sessions(status, started_at_ms DESC);
  CREATE INDEX idx_sessions_started_at ON imm_sessions(started_at_ms DESC);
  CREATE INDEX idx_sessions_ended_at ON imm_sessions(ended_at_ms DESC);
  CREATE INDEX idx_telemetry_session_sample ON imm_session_telemetry(session_id, sample_ms DESC);
  CREATE INDEX idx_telemetry_sample_ms ON imm_session_telemetry(sample_ms DESC);
  CREATE INDEX idx_events_session_ts ON imm_session_events(session_id, ts_ms DESC);
  CREATE INDEX idx_events_type_ts ON imm_session_events(event_type, ts_ms DESC);
  CREATE INDEX idx_rollups_day_video ON imm_daily_rollups(rollup_day, video_id);
  CREATE INDEX idx_rollups_month_video ON imm_monthly_rollups(rollup_month, video_id);
  CREATE INDEX idx_words_headword_word_reading ON imm_words(headword, word, reading);
  CREATE INDEX idx_words_frequency ON imm_words(frequency DESC);
  CREATE INDEX idx_kanji_kanji ON imm_kanji(kanji);
  CREATE INDEX idx_kanji_frequency ON imm_kanji(frequency DESC);
  CREATE INDEX idx_subtitle_lines_session_line ON imm_subtitle_lines(session_id, line_index);
  CREATE INDEX idx_subtitle_lines_video_line ON imm_subtitle_lines(video_id, line_index);
  CREATE INDEX idx_subtitle_lines_anime_line ON imm_subtitle_lines(anime_id, line_index);
  CREATE INDEX idx_word_line_occurrences_word ON imm_word_line_occurrences(word_id, line_id);
  CREATE INDEX idx_kanji_line_occurrences_kanji ON imm_kanji_line_occurrences(kanji_id, line_id);
  CREATE INDEX idx_media_art_cover_blob_hash ON imm_media_art(cover_blob_hash);
  CREATE INDEX idx_media_art_anilist_id ON imm_media_art(anilist_id);
  CREATE INDEX idx_media_art_cover_url ON imm_media_art(cover_url);
  CREATE INDEX idx_youtube_videos_channel_id ON imm_youtube_videos(channel_id);
  CREATE INDEX idx_youtube_videos_youtube_video_id ON imm_youtube_videos(youtube_video_id);
`;
