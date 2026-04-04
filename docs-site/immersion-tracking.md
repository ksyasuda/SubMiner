# Immersion Tracking

SubMiner can log your watching and mining activity to a local SQLite database, then surface it in the built-in stats dashboard. Tracking is enabled by default and can be turned off if you do not want local analytics.

When enabled, SubMiner records per-session statistics (watch time, subtitle lines seen, words encountered, cards mined) and maintains exact lifetime summary tables plus daily/monthly rollups. You can view that data in SubMiner's stats UI or query the database directly with any SQLite tool.

Episode completion for local `watched` state uses the shared `DEFAULT_MIN_WATCH_RATIO` (`85%`) value from `src/shared/watch-threshold.ts`.

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

## Stats Dashboard

The same immersion data powers the stats dashboard.

- In-app overlay: focus the visible overlay, then press the key from `stats.toggleKey` (default: `` ` `` / `Backquote`).
- Launcher command: run `subminer stats` to start the local stats server on demand and open the dashboard in your browser.
- Background server: run `subminer stats -b` to start or reuse a dedicated background stats daemon without keeping the launcher attached, and `subminer stats -s` to stop that daemon.
- Maintenance command: run `subminer stats cleanup` or `subminer stats cleanup -v` to backfill/repair vocabulary metadata (`headword`, `reading`, POS) and purge stale or excluded rows from `imm_words` on demand.
- Browser page: open `http://127.0.0.1:6969` directly if the local stats server is already running.

### Dashboard Tabs

#### Overview

Recent sessions, streak calendar, watch-time history, and a tracking snapshot with completed episodes/anime totals.

![Stats Overview](/screenshots/stats-overview.png)

#### Library

Cover-art library with search and sorting, per-series progress, episode drill-down, and direct links into mined cards.

When YouTube channel metadata is available, the Library tab groups videos by creator/channel and treats each tracked video as an episode-like entry inside that channel section.

![Stats Library](/screenshots/stats-library.png)

#### Trends

Watch time, sessions, words seen, and per-anime progress/pattern charts with configurable date ranges and grouping.

![Stats Trends](/screenshots/stats-trends.png)

#### Sessions

Expandable session history with new-word activity, cumulative totals, and pause/seek/card markers. Each session row exposes a hover-revealed ↗ button that navigates to the anime media-detail view for that session; pressing the back button there returns to the Sessions tab.

![Stats Sessions](/screenshots/stats-sessions.png)

#### Vocabulary

Top repeated words (click a bar to open the word), new-word timeline, frequency rank table with full readings, kanji breakdown, word exclusion list, and click-through occurrence drilldown with Mine Word / Mine Sentence / Mine Audio buttons.

![Stats Vocabulary](/screenshots/stats-vocabulary.png)

Stats server config lives under `stats`:

```jsonc
{
  "stats": {
    "toggleKey": "Backquote",
    "serverPort": 6969,
    "autoStartServer": true,
    "autoOpenBrowser": false
  }
}
```

- `toggleKey` is overlay-local, not a system-wide shortcut.
- `serverPort` controls the localhost dashboard URL.
- `autoStartServer` starts the local stats HTTP server on launch once immersion tracking is active, or reuses the dedicated background stats server when one is already running.
- `autoOpenBrowser` controls whether `subminer stats` launches the dashboard URL in your browser after ensuring the server is running.
- `subminer stats` forces the dashboard server to start even when `autoStartServer` is `false`.
- `subminer stats -b` starts or reuses the dedicated background stats daemon and exits after startup acknowledgement.
- The background stats daemon is separate from the normal SubMiner overlay app, so you can leave it running and still launch SubMiner later to watch or mine from video.
- `subminer stats -s` stops the dedicated background stats daemon without closing any browser tabs.
- `subminer stats` fails with an error when `immersionTracking.enabled` is `false`.
- `subminer stats cleanup` defaults to vocabulary cleanup, repairs stale `headword`, `reading`, and `part_of_speech` values, attempts best-effort MeCab backfill for legacy rows, and removes rows that still fail vocab filtering.

## Mining Cards from the Stats Page

The Vocabulary tab's word detail panel shows example lines from your viewing history. Each example line with a valid source file offers three mining buttons:

- **Mine Word** — performs a full Yomitan dictionary lookup for the word (definition, reading, pitch accent, etc.) via a short-lived hidden helper, then enriches the card with sentence audio, a screenshot or animated AVIF clip, the highlighted sentence, and metadata extracted from the source video file. Requires Anki and Yomitan dictionaries to be loaded.
- **Mine Sentence** — creates a sentence card directly with the `IsSentenceCard` flag set (for Lapis/Kiku workflows), along with audio, image, and translation from the secondary subtitle if available.
- **Mine Audio** — creates an audio-only card with the `IsAudioCard` flag, attaching only the sentence audio clip.

All three modes respect your `ankiConnect` config: deck, model, field mappings, media settings (static vs AVIF, quality, dimensions), audio padding, metadata pattern, and tags. Media generation runs in parallel for faster card creation.

Secondary subtitle text (typically English translations) is stored alongside primary subtitles during playback and used as the translation field when mining from the stats page.

### Word Exclusion List

The Vocabulary tab toolbar includes an **Exclusions** button for hiding words from all vocabulary views. Excluded words are stored in browser localStorage and can be managed (restored or cleared) from the exclusion modal. Exclusions affect stat cards, charts, the frequency rank table, and the word list.

## Retention Defaults

By default, SubMiner keeps all retention tables and raw data (`0` means keep all) while continuing daily/monthly rollup maintenance:

| Data type      | Retention |
| -------------- | --------- |
| Raw events     | 0 (keep all)   |
| Telemetry      | 0 (keep all)   |
| Sessions      | 0 (keep all)   |
| Daily rollups  | 0 (keep all)   |
| Monthly rollups | 0 (keep all) |

Maintenance runs on startup and every 24 hours. Vacuum runs only when `retention.vacuumIntervalDays` is non-zero.

In practice:

- Overview totals read from lifetime summary tables, so all-time watch time/cards/words stay exact even if raw query paths evolve.
- Anime and episode pages keep lifetime totals from summary tables while session drill-down still reads retained sessions directly. With the current defaults, both are kept forever.
- Trends can read the full available history because daily/monthly rollups are also kept forever by default.
- Vocabulary and kanji totals are cumulative and not bounded by the raw session retention knobs.

## Storage / Performance Model

The tracker is optimized for "keep everything" defaults:

- Exact all-time totals live in dedicated lifetime summary tables (`imm_lifetime_global`, `imm_lifetime_anime`, `imm_lifetime_media`).
- Ended-session totals are persisted onto `imm_sessions`, so most dashboard reads do not need to rescan raw telemetry.
- Daily and monthly rollups remain available for chart queries and coarse trend views.
- Subtitle text is stored once in `imm_subtitle_lines`; subtitle-line event payloads keep compact metadata only.
- Cover-art binaries are deduplicated through a shared blob store so episodes in the same series do not each carry duplicate image bytes.
- Hot tables have dedicated indexes for session time ranges, telemetry sample windows, frequency-ranked vocabulary, and cover-art lookup keys.

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
| `retention.sessionsDays` | Session retention |
| `retention.dailyRollupsDays` | Daily rollup retention |
| `retention.monthlyRollupsDays` | Monthly rollup retention |
| `retention.vacuumIntervalDays` | Minimum spacing between vacuums |
| `retentionMode` | `preset` or `advanced` |
| `retentionPreset` | `minimal`, `balanced`, or `deep-history` (used by `retentionMode`) |
| `lifetimeSummaries.global` | Maintain global lifetime totals |
| `lifetimeSummaries.anime` | Maintain per-anime lifetime totals |
| `lifetimeSummaries.media` | Maintain per-media lifetime totals |

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
  COALESCE(s.active_watched_ms, 0) AS active_watched_ms,
  COALESCE(s.words_seen, 0) AS words_seen,
  COALESCE(s.cards_mined, 0) AS cards_mined,
  CASE
    WHEN COALESCE(s.active_watched_ms, 0) > 0
      THEN COALESCE(s.words_seen, 0) / (COALESCE(s.active_watched_ms, 0) / 60000.0)
    ELSE NULL
  END AS words_per_min,
  CASE
    WHEN COALESCE(s.active_watched_ms, 0) > 0
      THEN (COALESCE(s.cards_mined, 0) * 60.0) / (COALESCE(s.active_watched_ms, 0) / 60000.0)
    ELSE NULL
  END AS cards_per_hour
FROM imm_sessions s
ORDER BY s.started_at_ms DESC
LIMIT ?;
```

### Lifetime anime totals

```sql
SELECT
  a.anime_id,
  a.canonical_title,
  la.total_sessions,
  la.total_active_ms,
  la.total_cards,
  la.total_words_seen,
  la.total_lines_seen,
  la.first_watched_ms,
  la.last_watched_ms
FROM imm_lifetime_anime la
JOIN imm_anime a ON a.anime_id = la.anime_id
ORDER BY la.last_watched_ms DESC
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
- SQLite tunings: `journal_mode=WAL`, `synchronous=NORMAL`, `foreign_keys=ON`, `busy_timeout=2500`, bounded WAL growth via `journal_size_limit`.
- Maintenance executes `PRAGMA optimize` after periodic cleanup.
- Rollups run incrementally from the last processed telemetry sample; startup performs a one-time bootstrap pass.
- Cover-art blobs are deduplicated into `imm_cover_art_blobs` and referenced from `imm_media_art`.
- Large-table reads are index-backed for `sample_ms`, session time windows, frequency-ranked words/kanji, and cover-art identity lookups.
- Workload-dependent tuning knobs remain at defaults unless you change them: `cache_size`, `mmap_size`, `temp_store`, `auto_vacuum`.

### Schema (v4)

Core tables:

- `imm_videos` — video key/title/source metadata
- `imm_sessions` — session UUID, video reference, timing/status, final denormalized totals
- `imm_session_telemetry` — high-frequency session aggregates over time
- `imm_session_events` — event stream with compact numeric event types
- `imm_subtitle_lines` — persisted subtitle text and timing per session/video

Lifetime summary tables:

- `imm_lifetime_global`
- `imm_lifetime_anime`
- `imm_lifetime_media`
- `imm_lifetime_applied_sessions`

Rollup tables:

- `imm_daily_rollups`
- `imm_monthly_rollups`

Vocabulary tables:

- `imm_words(id, headword, word, reading, first_seen, last_seen, frequency)`
- `imm_kanji(id, kanji, first_seen, last_seen, frequency)`

Media-art tables:

- `imm_media_art` — per-video cover metadata plus shared blob reference
- `imm_cover_art_blobs` — deduplicated image bytes keyed by blob hash
