---
id: TASK-28
title: Add SQLite-backed immersion tracking for mining sessions
status: Done
assignee: []
created_date: '2026-02-13 17:52'
updated_date: '2026-02-18 02:36'
labels:
  - analytics
  - backend
  - database
  - immersion
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
## Updated scope
Implement SQLite-first immersion tracking for mining sessions optimized for speed/size and designed for a future external DB adapter.

## Runtime defaults
- Flush batch policy: flush every `25` telemetry points or `500ms`.
- SQLite: `journal_mode = WAL`, `synchronous = NORMAL`, `foreign_keys = ON`, `busy_timeout = 2500ms`.
- Query target: `<150ms p95` for session/video/time-window reads at ~1M rows.
- In-memory queue cap: `1000` events; define explicit overflow behavior.
- Retention: events `7d`, telemetry `30d`, daily rollups `365d`, monthly rollups `5y`; prune on startup + every `24h`, vacuum on idle weekly.

## Concrete v1 schema (compact, size-aware)

```sql
CREATE TABLE imm_schema_version(
  schema_version INTEGER PRIMARY KEY,
  applied_at_ms INTEGER NOT NULL
);

CREATE TABLE imm_videos(
  video_id INTEGER PRIMARY KEY AUTOINCREMENT,
  video_key TEXT NOT NULL UNIQUE,
  canonical_title TEXT NOT NULL,
  source_type INTEGER NOT NULL,
  source_path TEXT,
  source_url TEXT,
  duration_ms INTEGER NOT NULL CHECK(duration_ms>=0),
  file_size_bytes INTEGER CHECK(file_size_bytes>=0),
  codec_id INTEGER, container_id INTEGER,
  width_px INTEGER, height_px INTEGER, fps_x100 INTEGER,
  bitrate_kbps INTEGER, audio_codec_id INTEGER,
  hash_sha256 TEXT, screenshot_path TEXT,
  metadata_json TEXT,
  created_at_ms INTEGER NOT NULL, updated_at_ms INTEGER NOT NULL
);

CREATE TABLE imm_sessions(
  session_id INTEGER PRIMARY KEY AUTOINCREMENT,
  session_uuid TEXT NOT NULL UNIQUE,
  video_id INTEGER NOT NULL,
  started_at_ms INTEGER NOT NULL, ended_at_ms INTEGER,
  status INTEGER NOT NULL,
  locale_id INTEGER, target_lang_id INTEGER,
  difficulty_tier INTEGER, subtitle_mode INTEGER,
  created_at_ms INTEGER NOT NULL, updated_at_ms INTEGER NOT NULL,
  FOREIGN KEY(video_id) REFERENCES imm_videos(video_id)
);

CREATE TABLE imm_session_telemetry(
  telemetry_id INTEGER PRIMARY KEY AUTOINCREMENT,
  session_id INTEGER NOT NULL,
  sample_ms INTEGER NOT NULL,
  total_watched_ms INTEGER NOT NULL DEFAULT 0,
  active_watched_ms INTEGER NOT NULL DEFAULT 0,
  lines_seen INTEGER NOT NULL DEFAULT 0,
  words_seen INTEGER NOT NULL DEFAULT 0,
  tokens_seen INTEGER NOT NULL DEFAULT 0,
  cards_mined INTEGER NOT NULL DEFAULT 0,
  lookup_count INTEGER NOT NULL DEFAULT 0,
  lookup_hits INTEGER NOT NULL DEFAULT 0,
  pause_count INTEGER NOT NULL DEFAULT 0,
  pause_ms INTEGER NOT NULL DEFAULT 0,
  seek_forward_count INTEGER NOT NULL DEFAULT 0,
  seek_backward_count INTEGER NOT NULL DEFAULT 0,
  media_buffer_events INTEGER NOT NULL DEFAULT 0,
  FOREIGN KEY(session_id) REFERENCES imm_sessions(session_id) ON DELETE CASCADE
);

CREATE TABLE imm_session_events(
  event_id INTEGER PRIMARY KEY AUTOINCREMENT,
  session_id INTEGER NOT NULL,
  ts_ms INTEGER NOT NULL,
  event_type INTEGER NOT NULL,
  line_index INTEGER,
  segment_start_ms INTEGER,
  segment_end_ms INTEGER,
  words_delta INTEGER NOT NULL DEFAULT 0,
  cards_delta INTEGER NOT NULL DEFAULT 0,
  payload_json TEXT,
  FOREIGN KEY(session_id) REFERENCES imm_sessions(session_id) ON DELETE CASCADE
);

CREATE TABLE imm_daily_rollups(
  rollup_day INTEGER NOT NULL,
  video_id INTEGER,
  total_sessions INTEGER NOT NULL DEFAULT 0,
  total_active_min REAL NOT NULL DEFAULT 0,
  total_lines_seen INTEGER NOT NULL DEFAULT 0,
  total_words_seen INTEGER NOT NULL DEFAULT 0,
  total_tokens_seen INTEGER NOT NULL DEFAULT 0,
  total_cards INTEGER NOT NULL DEFAULT 0,
  cards_per_hour REAL,
  words_per_min REAL,
  lookup_hit_rate REAL,
  PRIMARY KEY (rollup_day, video_id)
);

CREATE TABLE imm_monthly_rollups(
  rollup_month INTEGER NOT NULL,
  video_id INTEGER,
  total_sessions INTEGER NOT NULL DEFAULT 0,
  total_active_min REAL NOT NULL DEFAULT 0,
  total_lines_seen INTEGER NOT NULL DEFAULT 0,
  total_words_seen INTEGER NOT NULL DEFAULT 0,
  total_tokens_seen INTEGER NOT NULL DEFAULT 0,
  total_cards INTEGER NOT NULL DEFAULT 0,
  PRIMARY KEY (rollup_month, video_id)
);

CREATE INDEX idx_sessions_video_started ON imm_sessions(video_id, started_at_ms DESC);
CREATE INDEX idx_sessions_status_started ON imm_sessions(status, started_at_ms DESC);
CREATE INDEX idx_telemetry_session_sample ON imm_session_telemetry(session_id, sample_ms DESC);
CREATE INDEX idx_events_session_ts ON imm_session_events(session_id, ts_ms DESC);
CREATE INDEX idx_events_type_ts ON imm_session_events(event_type, ts_ms DESC);
CREATE INDEX idx_rollups_day_video ON imm_daily_rollups(rollup_day, video_id);
CREATE INDEX idx_rollups_month_video ON imm_monthly_rollups(rollup_month, video_id);
```

Notes
- Integer enums keep hot rows narrow and fast; resolve labels in app layer.
- JSON fields are only for non-core overflow attributes (bounded by policy).
- `source_type`/`status`/`subtitle_mode` are compact enums and keep strings out of high-frequency rows.
- This schema is a v1 contract for TASK-32 adapter work later.

## Future portability
- Keep all analytics logic behind a storage interface in TASK-32.
- Keep raw SQL/DDL details inside adapters.

## Execution principle
- Tracking is strictly asynchronous and must never block hot paths.
- Tokenization/rendering pipelines must not await DB operations.
- If tracker queue is saturated, user experience must remain unchanged; telemetry may be dropped with bounded loss and explicit internal warning/logging.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 A SQLite database schema is defined and created automatically (or initialized on startup) for immersion tracking if not present.
- [x] #2 Recorded events persist at least the following fields per session/item: video name, video directory/URL, video length, lines seen, words/tokens seen, cards mined.
- [x] #3 Tracking defaults to storing data in SQLite without requiring additional DB setup for local usage.
- [x] #4 Additional extractable metadata from video files is captured and stored when available (e.g., dimensions, duration, codec, fps, file size/hash, optional screenshot path).
- [x] #5 Tracking does not degrade mining throughput and handles duplicate/missing metadata fields safely.
- [x] #6 Query/read paths exist to support future richer statistics generation (e.g., totals by video, throughput, quality metrics).
- [x] #7 Schema design and implementation include clear migration/versioning strategy for future fields.
- [x] #8 Schema uses compact numeric/tiny integer types where practical and minimizes repeated TEXT payloads to balance write/read speed and file size.
- [x] #9 High-frequency writes are batched (or buffered) with periodic checkpoints so writes do not fsync per telemetry point.
- [x] #10 Event retention and rollup strategy is documented: raw event retention, summary tables, and compaction policy to bound DB size.
- [x] #11 Query performance targets are addressed with index strategy and a documented plan for index coverage (session-by-video, time-window, event-type, card/count lookups).
- [x] #12 Migration/versioning strategy supports future backend portability without requiring analytics-layer rewrite (schema version table + adapter boundary specified).
- [x] #13 Task defines operational defaults: flush every 25 events or 500ms, WAL+NORMAL, queue cap of 1000 rows, in-flight payload cap of 256B, and explicit overflow behavior.
- [x] #14 Task defines retention defaults and maintenance cadence: events 7d, telemetry 30d, daily 365d, monthly 5y, startup + 24h prune and idle-weekly vacuum.
- [x] #15 Task documents expected query performance target (150ms p95) and storage growth guardrails for typical local usage up to ~1M events.
- [x] #16 #13 Concrete DDL (tables + indexes + pragmas) is captured in task docs and used as implementation reference.
- [x] #17 #14 v1 retention policy, batch policy, and maintenance schedule are explicitly implemented and configurable.
- [x] #18 #15 Query templates for timeline/throughput/rollups are defined in implementation docs.
- [x] #19 #16 Queue cap, payload cap, and overflow behavior are implemented and documented.
- [x] #20 #20 All tracking writes are strictly asynchronous and non-blocking from tokenization/render loops; hot paths must never await persistence.
- [x] #21 #21 Queue saturation handling is explicit: bounded queue with deterministic policy (drop oldest, drop newest, or backpressure) and no impact on on-screen token colorization or line rendering.
- [x] #22 #22 Tracker failures/timeouts are swallowed from hot path with optional background retry and failure counters/logging for observability.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Progress review (2026-02-17): `src/core/services/immersion-tracker-service.ts` now implements SQLite-first schema init, WAL/NORMAL pragmas, async queue + batch flush (25/500ms), queue cap 1000 with drop-oldest overflow policy, payload clamp (256B), retention pruning (events 7d, telemetry 30d, daily 365d, monthly 5y), startup+24h maintenance, weekly vacuum, rollup maintenance, and query paths (`getSessionSummaries`, `getSessionTimeline`, `getDailyRollups`, `getMonthlyRollups`, `getQueryHints`).

Metadata capture is implemented for local media via ffprobe/stat/SHA-256 (`captureVideoMetadataAsync`, `getLocalVideoMetadata`) with safe null handling for missing fields.

Remaining scope before close: AC #17 and #18 are still open. Current retention/batch defaults are hardcoded constants (implemented but not externally configurable), and there is no dedicated implementation doc section defining query templates for timeline/throughput/rollups outside code.

Tests present in `src/core/services/immersion-tracker-service.test.ts` validate session UUIDs, session finalization telemetry persistence, monthly rollups, and prepared statement reuse; broader retrievability coverage may still be expanded later if desired.

Completed remaining scope (2026-02-18): retention/batch/maintenance defaults are now externally configurable under `immersionTracking` (`batchSize`, `flushIntervalMs`, `queueCap`, `payloadCapBytes`, `maintenanceIntervalMs`, and nested `retention.*` day windows). Runtime wiring now passes config policy into `ImmersionTrackerService` and service applies bounded values with safe fallbacks.

Implementation docs now include query templates and storage behavior in `docs/immersion-tracking.md` (timeline, throughput summary, daily/monthly rollups), plus config reference updates in `docs/configuration.md` and examples.

Validation/tests expanded: `src/config/config.test.ts` now covers immersion tuning parse+fallback warnings; `src/core/services/immersion-tracker-service.test.ts` adds minimum persisted/retrievable field checks and configurable policy checks.

Verification run: `pnpm run build && node --test dist/config/config.test.js dist/core/services/immersion-tracker-service.test.js` passed; sqlite-specific tracker tests are skipped automatically in environments without `node:sqlite` support.
<!-- SECTION:NOTES:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 SQLite tracking table(s), migration history table, and indices created as part of startup or init path.
- [x] #2 Unit/integration coverage (or validated test plan) confirms minimum fields are persisted and retrievable.
- [x] #3 README or docs updated with storage schema, retention defaults, and extension points.
- [x] #4 Migration and retention defaults are documented (pruning frequency, rollup cadence, expected disk growth profile).
- [x] #5 Performance-safe write path behavior is documented (batch commit interval/size, WAL mode, sync mode).
- [x] #6 A follow-up ticket captures and tracks non-SQLite backend abstraction work.
- [x] #7 The implementation doc includes the exact schema, migration version, and index set.
- [x] #8 Performance-size tradeoffs are clearly documented (batching, enum columns, bounded JSON, TTL retention).
- [x] #9 Rollup/retention behavior is in place with explicit defaults and cleanup cadence.
<!-- DOD:END -->
