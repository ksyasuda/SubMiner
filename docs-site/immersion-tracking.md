# Immersion Tracking

SubMiner can log your watching and mining activity to a local SQLite database. This is optional and disabled by default.

When enabled, SubMiner records per-session statistics (watch time, subtitle lines seen, words encountered, cards mined) and maintains daily and monthly rollups. You can query the database directly with any SQLite tool to track your progress over time.

## Enabling

```jsonc
{
  "immersionTracking": {
    "enabled": true,
    "dbPath": ""
  }
}
```

- Leave `dbPath` empty to use the default location (`immersion.sqlite` in SubMiner's app-data directory).
- Set an explicit path to move the database (useful for backups, cloud syncing, or external tools).

## Retention Defaults

Data is kept for the following durations before automatic cleanup:

| Data type      | Retention |
| -------------- | --------- |
| Raw events     | 7 days    |
| Telemetry      | 30 days   |
| Daily rollups  | 1 year    |
| Monthly rollups | 5 years  |

Maintenance runs on startup and every 24 hours. Vacuum runs weekly.

## Configurable Knobs

All policy options live under `immersionTracking` in your config:

| Option | Description |
| ------ | ----------- |
| `batchSize` | Writes per flush batch |
| `flushIntervalMs` | Max delay between flushes (default: 500ms) |
| `queueCap` | Max queued writes before oldest are dropped |
| `payloadCapBytes` | Max payload size per write |
| `maintenanceIntervalMs` | How often maintenance runs |
| `retention.eventsDays` | Raw event retention |
| `retention.telemetryDays` | Telemetry retention |
| `retention.dailyRollupsDays` | Daily rollup retention |
| `retention.monthlyRollupsDays` | Monthly rollup retention |
| `retention.vacuumIntervalDays` | Minimum spacing between vacuums |

## Query Templates

### Session timeline

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
ORDER BY sample_ms DESC, telemetry_id DESC
LIMIT ?;
```

### Session throughput summary

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

### Daily rollups

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

### Monthly rollups

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

## Technical Details

- Write path is asynchronous and queue-backed. Hot paths (subtitle parsing, render, token flows) enqueue telemetry and never await SQLite writes.
- Queue overflow policy: drop oldest queued writes, keep newest.
- SQLite pragmas: `journal_mode=WAL`, `synchronous=NORMAL`, `foreign_keys=ON`, `busy_timeout=2500`.
- Rollups run incrementally from the last processed telemetry sample; startup performs a one-time bootstrap pass.
- If retention pruning removes telemetry/session rows, maintenance triggers a full rollup rebuild to resync historical aggregates.

### Schema (v3)

Core tables:

- `imm_videos` — video key/title/source metadata
- `imm_sessions` — session UUID, video reference, timing/status
- `imm_session_telemetry` — high-frequency session aggregates over time
- `imm_session_events` — event stream with compact numeric event types

Rollup tables:

- `imm_daily_rollups`
- `imm_monthly_rollups`

Vocabulary tables:

- `imm_words(id, headword, word, reading, first_seen, last_seen, frequency)`
- `imm_kanji(id, kanji, first_seen, last_seen, frequency)`
