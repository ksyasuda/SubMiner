# Immersion Tracking

SubMiner can log your watching and mining activity to a local SQLite database, then surface it in the built-in stats dashboard. Tracking is enabled by default and can be turned off if you do not want local analytics.

"Immersion" here means time spent watching and reading native Japanese content. **All data stays on your computer** - nothing is uploaded anywhere. (SQLite is just a single-file database; you do not need to install or manage anything.)

When enabled, SubMiner records per-session statistics (watch time, subtitle lines seen, words encountered, cards mined) and maintains exact lifetime summary tables plus daily/monthly rollups. You can view that data in SubMiner's stats UI or query the database directly with any SQLite tool.

::: tip For most users
Just leave tracking on and use the built-in [Stats Dashboard](#stats-dashboard). The retention, performance, SQL, and schema sections further down are reference material for advanced users who want to inspect or tune the database - you can safely skip them.
:::

Episode completion for local `watched` state uses the shared `DEFAULT_MIN_WATCH_RATIO` (`85%`) value from `src/shared/watch-threshold.ts`.

## Enabling

```jsonc
{
  "immersionTracking": {
    "enabled": true,
    "dbPath": "",
  },
}
```

- Leave `dbPath` empty to use the default location (`immersion.sqlite` in SubMiner's app-data directory).
- Set an explicit path to move the database (useful for backups, cloud syncing, or external tools).
- To share stats and watch history between two machines, use [`subminer sync <host>`](/launcher-script#sync-between-machines) instead of file-level cloud sync — it merges both databases without one side overwriting the other.

## Stats Dashboard

The same immersion data powers the stats dashboard.

- In-app overlay: focus the visible overlay, then press the key from `stats.toggleKey` (default: `` ` `` / `Backquote`).
- Launcher command: run `subminer stats` to start the local stats server on demand (it also opens the dashboard in your browser when `stats.autoOpenBrowser` is enabled; the default is `false`).
- Background server: run `subminer stats -b` to start or reuse a dedicated background stats daemon without keeping the launcher attached, and `subminer stats -s` to stop that daemon.
- Maintenance commands: run `subminer stats cleanup` or `subminer stats cleanup -v` to backfill/repair vocabulary metadata (`headword`, `reading`, POS) and purge stale or excluded rows from `imm_words` on demand; `subminer stats cleanup -l` repairs lifetime summary tables non-destructively (recomputed from per-episode history, so lifetime totals older than the session retention window are kept); `subminer stats cleanup --duplicate-lines` collapses repeated lines left behind by typeset subtitles (see [Repeated Line Cleanup](#repeated-line-cleanup)). `subminer stats rebuild` and `subminer stats backfill` rebuild or backfill rollup data.
- Browser page: open `http://127.0.0.1:6969` directly if the local stats server is already running.

### Dashboard Tabs

#### Overview

Recent sessions, streak calendar, watch-time history, and a tracking snapshot with completed episodes/anime totals.

![Stats Overview](/screenshots/stats-overview.png)

#### Library

Cover-art library with search and sorting, per-series progress, episode drill-down, and direct links into mined cards.

Local files and Jellyfin items with detected season numbers are split into season-specific library entries, so `Season 1` and `Season 2` folders do not merge into one show card.

When older stats already grouped multiple seasons under one series entry, SubMiner moves parsed episodes into the season-specific entries on startup and rebuilds the affected summaries.

Jellyfin stream URLs are normalized to stable item links before stats titles are shown, so playback query parameters are not displayed in the dashboard.

When YouTube channel metadata is available, the Library tab groups videos by creator/channel and treats each tracked video as an episode-like entry inside that channel section.

A library entry is identified by its parsed title plus any detected season, so the same show can end up on several cards when releases disagree about the title or omit the season tag. Two fixes are available:

- **Merge duplicates.** Hit **Select** above the grid, tick the cards that are the same show, and choose **Merge Selected**. Pick which entry to keep in the dialog; every episode moves onto it and the other cards are removed. Nothing is deleted, so sessions, mined cards and watch time all carry over. SubMiner remembers the merged title variants, so future episodes parsed with one of those names join the kept entry instead of recreating a duplicate card.
- **Move a single episode.** Hover an episode row in a title's episode list and use the **→** button to reassign it to another library entry. The correction is remembered, so later filename parsing or Jellyfin metadata cannot move that episode back. For local files, later episodes in the same directory inherit the correction when their detected seasons are compatible and every manual correction there points to the same entry; a file that parses to a title which already has its own library entry keeps that identity instead. Conflicting seasons or manual destinations are left for review. If the move empties the old entry, that card is removed and you are returned to the grid.

Once cover art resolves a series to an AniList entry, cards with compatible seasons are folded together automatically only when the searched title exactly matches an AniList title or synonym. A fuzzy result that points at an AniList entry already used by another card appears as a **Possible duplicate** review above the Library grid instead. Choose **Review merge** to compare the cards and pick which one to keep, or **Not duplicates** to dismiss that suggestion permanently. Entries with conflicting explicit season numbers are left alone rather than merged or suggested.

Open a title and use **Delete Entry** in its header to remove a mistakenly tracked show outright. This deletes every episode of that title along with their sessions, subtitle lines, rollups and cover art, drops the words and kanji that were only seen there, and removes the card from the Library grid. Individual episodes and sessions can still be deleted on their own from the episode list and session rows. Entry deletion is refused while that title is the one currently playing.

![Stats Library](/screenshots/stats-library.png)

#### Trends

Grouped into Activity (per-day/month watch time, cards, words, sessions), Cumulative Totals (running totals incl. new words seen and episodes), Efficiency (words/min, cards/hour, lookups per 100 words), Patterns (watch time by day of week and hour), and per-anime Library charts — all with configurable date ranges and grouping.

![Stats Trends](/screenshots/stats-trends.png)

#### Sessions

Expandable session history with new-word activity, cumulative totals, and pause/seek/card markers. Each session row exposes a hover-revealed ↗ button that navigates to the anime media-detail view for that session; pressing the back button there returns to the Sessions tab.

![Stats Sessions](/screenshots/stats-sessions.png)

#### Vocabulary

Top repeated words (click a bar to open the word), new-word timeline, cross-title and frequency rank tables with Hide Known / Hide Kana filters, kanji breakdown, word exclusion list, and click-through occurrence drilldown with Mine Word / Mine Sentence / Mine Audio buttons.

![Stats Vocabulary](/screenshots/stats-vocabulary.png)

#### Search

Realtime search across tracked primary subtitle lines and media titles. Results show the source media, session, line number, timing, and sentence text. Secondary subtitle text is not shown or searched here because separate subtitle tracks may not line up sentence-for-sentence. Sentence cards can be mined from any result with a valid local source and timing. Word and audio card buttons appear only when the searched word exactly appears in the primary sentence text; matching text is highlighted in the result.

Stats server config lives under `stats`:

```jsonc
{
  "stats": {
    "toggleKey": "Backquote",
    "markWatchedKey": "KeyW",
    "serverPort": 6969,
    "autoStartServer": true,
    "autoOpenBrowser": false,
  },
}
```

- `toggleKey` is overlay-local, not a system-wide shortcut.
- `markWatchedKey` toggles the watched state of the highlighted entry inside the stats dashboard.
- `serverPort` controls the localhost dashboard URL.
- `autoStartServer` starts the local stats HTTP server on launch once immersion tracking is active, or reuses the dedicated background stats server when one is already running. Background app launches (`subminer app`) start the stats server immediately, registering it so later launches reuse it instead of starting another one.
- `autoOpenBrowser` controls whether `subminer stats` launches the dashboard URL in your browser after ensuring the server is running.
- `subminer stats` forces the dashboard server to start even when `autoStartServer` is `false`.
- `subminer stats -b` starts or reuses the dedicated background stats daemon and exits after startup acknowledgement.
- The background stats daemon is separate from the normal SubMiner overlay app, so you can leave it running and still launch SubMiner later to watch or mine from video.
- `subminer stats -s` stops the dedicated background stats daemon without closing any browser tabs.
- `subminer stats` fails with an error when `immersionTracking.enabled` is `false`.
- `subminer stats cleanup` defaults to vocabulary cleanup, repairs stale `headword`, `reading`, and `part_of_speech` values, attempts best-effort MeCab backfill for legacy rows, and removes rows that still fail vocab filtering.

## Mining Cards from the Stats Page

The Search tab and the Vocabulary tab's word detail panel both mine from subtitle lines in your viewing history. Search matches sentence text and media titles, and **Search by headword** is enabled by default so dictionary-form searches such as `知らない` can find tracked subtitle lines with inflected variants. Turn that toggle off for exact text/title matching only. Each line with a valid source file offers sentence-card mining; word/audio mining is available when the selected word or searched word appears in the sentence:

- **Mine Word** - performs a full Yomitan dictionary lookup for the word (definition, reading, pitch accent, etc.) via a short-lived hidden helper, then enriches the card with sentence audio, a screenshot or animated AVIF clip, the highlighted sentence, and metadata extracted from the source video file. Requires Anki and Yomitan dictionaries to be loaded.
- **Mine Sentence** - creates a sentence card directly with the `IsSentenceCard` flag set (for Lapis/Kiku workflows), along with audio and image from the source video.
- **Mine Audio** - creates an audio-only card with the `IsAudioCard` flag, attaching only the sentence audio clip.

All three modes respect your `ankiConnect` config: deck, model, field mappings, media settings (static vs AVIF, quality, dimensions), audio padding, metadata pattern, and tags. Media generation runs in parallel for faster card creation.

Secondary subtitle text (typically English translations) is stored alongside primary subtitles during playback and can be used as the translation field when mining sentence cards from Search or vocabulary occurrences. The Search tab does not use that text for display or matching.

### Word Exclusion List

The Vocabulary tab toolbar includes an **Exclusions** button for hiding words from all vocabulary views. Excluded words are stored in the immersion database, with older browser localStorage exclusions imported on first load after upgrade. They can be managed (restored or cleared) from the exclusion modal. Exclusions affect stat cards, charts, the frequency rank table, and the word list.

### Repeated Line Cleanup

Karaoke openings and animated signs are authored as one subtitle event per animation frame, all carrying the same text. Playback reports every one of those frames, so a single OP lyric could be recorded hundreds of times and dominate "Top Repeated Words".

Recording now collapses those runs as they happen, matching what the subtitle sidebar shows:

- When a typeset ASS file stores a clean lyric or sign in a timed authoring comment, or in full-line events surrounding generated fragments, the matching complete line is recorded once. The repeated glyph or clip-animation frames are not recorded. Dialogue spoken while such an animation is on screen records as itself, without the fragment lines beside it.
- When karaoke styling redraws the same complete lyric across consecutive color or highlight phases, those phases are combined into one line with their full timing. Repeated ordinary dialogue remains separate.
- When the active subtitle source has been parsed, its cue list has already had duplicate events and animation bursts merged. A line landing inside a surviving cue but after that cue's start is a frame the sidebar merged away, and is not recorded.
- When no parsed cue covers the live timing, including while a subtitle source is changing or shifted, the strict metadata-free rule applies: a run of identical, contiguous lines each shorter than 0.1s stops being recorded after a few frames. Runs are tracked per line of text, so dual-line karaoke (a kanji and a romaji line frame-flipped together) collapses both lines. Ordinary repeated dialogue, and lines held for a normal beat, always record.

For stats recorded before this, the Vocabulary tab toolbar has a **Duplicates** button:

- Pick how far back to look (7 days, 30 days, 90 days, 1 year, or all time). A narrower window does less work and keeps older history untouched.
- **Scan** reports the bursts found, the lines they added, and the word and kanji counts they inflated, without writing anything.
- **Clean Up** applies exactly what the scan reported: each run collapses to its first line (extended to cover the run), and the removed lines' word and kanji occurrences are subtracted from the vocabulary aggregates.

The same thing runs from the terminal:

```bash
subminer stats cleanup --duplicate-lines --dry-run --lookback-days 30
subminer stats cleanup --duplicate-lines --lookback-days 30
```

`--duplicate-lines` (short: `-d`) picks the cleanup mode, so it cannot be combined with `--vocab` or `--lifetime`, and `--dry-run` and `--lookback-days <days>` only apply to it. Omitting `--lookback-days` scans all history; the value must be at least one day.

The cleanup chains runs per line of text, so interleaved dual-line karaoke collapses each of its lines. It also removes the short residue the live rule stores before a run is long enough to recognize: a run one frame short of the usual minimum qualifies when every event is under the strict 0.1s bound.

Runs never cross a session boundary, so rewatching an episode keeps both watches. Session telemetry (watch time, lines seen, tokens seen) and the rollups derived from it are left as recorded: they are cumulative samples taken during playback, and cannot be recomputed for sessions whose raw rows have since been pruned.

## Retention Defaults

By default, SubMiner keeps all retention tables and raw data (`0` means keep all) while continuing daily/monthly rollup maintenance:

| Data type       | Retention    |
| --------------- | ------------ |
| Raw events      | 0 (keep all) |
| Telemetry       | 0 (keep all) |
| Sessions        | 0 (keep all) |
| Daily rollups   | 0 (keep all) |
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

| Option                         | Description                                                        |
| ------------------------------ | ------------------------------------------------------------------ |
| `batchSize`                    | Writes per flush batch                                             |
| `flushIntervalMs`              | Max delay between flushes (default: 500ms)                         |
| `queueCap`                     | Max queued writes before oldest are dropped                        |
| `payloadCapBytes`              | Max payload size per write                                         |
| `maintenanceIntervalMs`        | How often maintenance runs                                         |
| `retention.eventsDays`         | Raw event retention                                                |
| `retention.telemetryDays`      | Telemetry retention                                                |
| `retention.sessionsDays`       | Session retention                                                  |
| `retention.dailyRollupsDays`   | Daily rollup retention                                             |
| `retention.monthlyRollupsDays` | Monthly rollup retention                                           |
| `retention.vacuumIntervalDays` | Minimum spacing between vacuums                                    |
| `retentionMode`                | `preset` or `advanced`                                             |
| `retentionPreset`              | `minimal`, `balanced`, or `deep-history` (used by `retentionMode`) |
| `lifetimeSummaries.global`     | Maintain global lifetime totals                                    |
| `lifetimeSummaries.anime`      | Maintain per-anime lifetime totals                                 |
| `lifetimeSummaries.media`      | Maintain per-media lifetime totals                                 |

## Query Templates

### Session timeline

```sql
SELECT
  sample_ms,
  total_watched_ms,
  active_watched_ms,
  lines_seen,
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
  COALESCE(s.tokens_seen, 0) AS tokens_seen,
  COALESCE(s.cards_mined, 0) AS cards_mined,
  CASE
    WHEN COALESCE(s.active_watched_ms, 0) > 0
      THEN COALESCE(s.tokens_seen, 0) / (COALESCE(s.active_watched_ms, 0) / 60000.0)
    ELSE NULL
  END AS tokens_per_min,
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
  la.total_tokens_seen,
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
  total_tokens_seen,
  total_cards,
  cards_per_hour,
  tokens_per_min,
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

### Schema (v18)

The exact schema version lives in `SCHEMA_VERSION` (`src/core/services/immersion-tracker/types.ts`) and is recorded in the `imm_schema_version` table.

Core tables:

- `imm_videos` - video key/title/source metadata
- `imm_anime` - anime/series metadata referenced by videos and lifetime tables
- `imm_sessions` - session UUID, video reference, timing/status, final denormalized totals
- `imm_session_telemetry` - high-frequency session aggregates over time
- `imm_session_events` - event stream with compact numeric event types
- `imm_subtitle_lines` - persisted subtitle text and timing per session/video
- `imm_youtube_videos` - YouTube video/channel metadata for tracked videos

Lifetime summary tables:

- `imm_lifetime_global`
- `imm_lifetime_anime`
- `imm_lifetime_media`
- `imm_lifetime_applied_sessions`

Rollup tables:

- `imm_daily_rollups`
- `imm_monthly_rollups`
- `imm_rollup_state` - incremental rollup progress bookkeeping

Vocabulary tables:

- `imm_words(id, headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency, frequency_rank)` with `UNIQUE(headword, word, reading)`
- `imm_kanji(id, kanji, first_seen, last_seen, frequency)`
- `imm_word_line_occurrences` / `imm_kanji_line_occurrences` - word/kanji ↔ subtitle-line occurrence links
- `imm_stats_excluded_words` - vocabulary exclusion list managed from the dashboard

Media-art tables:

- `imm_media_art` - per-video cover metadata plus shared blob reference
- `imm_cover_art_blobs` - deduplicated image bytes keyed by blob hash
