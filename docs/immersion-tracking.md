# Immersion Tracking Storage

SubMiner stores immersion analytics in local SQLite (`immersion.sqlite`) by default.

## Runtime Model

- Write path is asynchronous and queue-backed.
- Hot paths (subtitle parsing/render/token flows) enqueue telemetry/events and never await SQLite writes.
- Background line processing also upserts to `imm_words` and `imm_kanji`.
- Queue overflow policy is deterministic: drop oldest queued writes, keep newest.
- Flush policy defaults to `25` writes or `500ms` max delay.
- SQLite pragmas: `journal_mode=WAL`, `synchronous=NORMAL`, `foreign_keys=ON`, `busy_timeout=2500`.
- Rollups now run incrementally from the last processed telemetry sample; startup performs a one-time bootstrap rebuild-equivalent pass.
- If retention pruning removes telemetry/session rows, maintenance triggers a full rollup rebuild to resync historical aggregates.

## Schema (v3)

Schema versioning table:

- `imm_schema_version(schema_version PK, applied_at_ms)`

Core entities:

- `imm_videos`: video key/title/source metadata + optional media metadata fields, `CREATED_DATE`/`LAST_UPDATE_DATE`
- `imm_sessions`: session UUID, video reference, timing/status fields, `CREATED_DATE`/`LAST_UPDATE_DATE`
- `imm_session_telemetry`: high-frequency session aggregates over time, `CREATED_DATE`/`LAST_UPDATE_DATE`
- `imm_session_events`: event stream with compact numeric event types, `CREATED_DATE`/`LAST_UPDATE_DATE`

Rollups:

- `imm_daily_rollups`: includes `CREATED_DATE`/`LAST_UPDATE_DATE`
- `imm_monthly_rollups`: includes `CREATED_DATE`/`LAST_UPDATE_DATE`

Vocabulary:

- `imm_words(id, headword, word, reading, first_seen, last_seen, frequency)`
- `imm_kanji(id, kanji, first_seen, last_seen, frequency)`
- `first_seen`/`last_seen` store Unix timestamps and are upserted with line ingestion

Primary index coverage:

- session-by-video/time: `idx_sessions_video_started`
- session-by-status/time: `idx_sessions_status_started`
- timeline reads: `idx_telemetry_session_sample`
- event timeline/type reads: `idx_events_session_ts`, `idx_events_type_ts`
- rollup reads: `idx_rollups_day_video`, `idx_rollups_month_video`

## Retention and Maintenance Defaults

- Raw events: `7d`
- Telemetry: `30d`
- Daily rollups: `365d`
- Monthly rollups: `5y`
- Maintenance cadence: startup + every `24h`
- Vacuum cadence: idle weekly (`7d` minimum spacing)

Retention cleanup and rollup refresh stay in service maintenance orchestration + `src/core/services/immersion-tracker/maintenance.ts`.

## Configurable Policy Knobs

All knobs are under `immersionTracking` in config:

- `batchSize`
- `flushIntervalMs`
- `queueCap`
- `payloadCapBytes`
- `maintenanceIntervalMs`
- `retention.eventsDays`
- `retention.telemetryDays`
- `retention.dailyRollupsDays`
- `retention.monthlyRollupsDays`
- `retention.vacuumIntervalDays`

These map directly to runtime tracker policy and allow tuning without code changes.

## Query Templates

Timeline for one session:

```sql
SELECT
  sample_ms,
  total_watched_ms,
  active_watched_ms,
  lines_seen,
  words_seen,
  tokens_seen,
  cards_mined
FROM imm_session_telemetry
WHERE session_id = ?
ORDER BY sample_ms DESC
LIMIT ?;
```

Session throughput summary:

```sql
SELECT
  s.session_id,
  s.video_id,
  s.started_at_ms,
  s.ended_at_ms,
  COALESCE(SUM(t.active_watched_ms), 0) AS active_watched_ms,
  COALESCE(SUM(t.words_seen), 0) AS words_seen,
  COALESCE(SUM(t.cards_mined), 0) AS cards_mined,
  CASE
    WHEN COALESCE(SUM(t.active_watched_ms), 0) > 0
      THEN COALESCE(SUM(t.words_seen), 0) / (COALESCE(SUM(t.active_watched_ms), 0) / 60000.0)
    ELSE NULL
  END AS words_per_min,
  CASE
    WHEN COALESCE(SUM(t.active_watched_ms), 0) > 0
      THEN (COALESCE(SUM(t.cards_mined), 0) * 60.0) / (COALESCE(SUM(t.active_watched_ms), 0) / 60000.0)
    ELSE NULL
  END AS cards_per_hour
FROM imm_sessions s
LEFT JOIN imm_session_telemetry t ON t.session_id = s.session_id
GROUP BY s.session_id
ORDER BY s.started_at_ms DESC
LIMIT ?;
```

Daily rollups:

```sql
SELECT
  rollup_day,
  video_id,
  total_sessions,
  total_active_min,
  total_lines_seen,
  total_words_seen,
  total_tokens_seen,
  total_cards,
  cards_per_hour,
  words_per_min,
  lookup_hit_rate
FROM imm_daily_rollups
ORDER BY rollup_day DESC, video_id DESC
LIMIT ?;
```

Monthly rollups:

```sql
SELECT
  rollup_month,
  video_id,
  total_sessions,
  total_active_min,
  total_lines_seen,
  total_words_seen,
  total_tokens_seen,
  total_cards
FROM imm_monthly_rollups
ORDER BY rollup_month DESC, video_id DESC
LIMIT ?;
```
