# Stats Trends Data Flow

read_when: touching stats trend charts, changing stats API payloads, or debugging dashboard performance

## Summary

Trend charts now consume one chart-oriented backend payload from `/api/stats/trends/dashboard`.

## Why

- remove repeated client-side dataset rebuilding in `TrendsTab`
- collapse multiple network round-trips into one request
- keep heavy chart shaping close to tracker/query logic

## Date Ranges And Buckets

- A range resolves to calendar days (`resolveTrendWindow`). Rollups are selected by day span, never by "last N rows", so sparse activity cannot pull older days into a range. Custom ranges are `{ startDay, endDay }` from `range=custom&fromDay&toDay`.
- `all` spans the recorded data, not today.
- Day and week buckets read daily rollups. Month buckets read monthly rollups for months wholly in range, because the `minimal` and `deep-history` retention presets prune daily rollups first, and daily rollups for a partly covered month. Weeks start on Monday.
- `buckets` is the one axis for `activity`, `progress`, and `ratios`, index for index. Charts use it to map a clicked bar back to its days.
- Each `progress` point is the total at the end of its period, and the chart adds the zero starting point.

## Data Sources

- rollup-backed:
  - activity charts and running totals
  - `titleTimeline`: sparse per-title minutes for active days only, because a dense title x day payload grew to several MB
- session-metric-backed:
  - lookups and lookup rate
  - `watchTimeByDayHour`, attributed to the hour each session started
- rollup word counts and session metrics share one session query per request (`getTrendSessionRows`), so the per-session display-word count runs once instead of once per rollup query
- `imm_videos.watched`-backed:
  - finished episodes count each watched video once, on the last day it has a rollup, since the flag has no timestamp
- vocabulary-backed:
  - `getNewWordFirstSeenRows` folds `imm_words` rows into dictionary words first, because the table is unique on (headword, word, reading) and stores each conjugation separately
  - the aggregate scans every word, so it is cached per connection until `total_changes()` or `PRAGMA data_version` moves
  - Trends new-word metrics exclude names; the Vocabulary tab offers both
  - `imm_lexical_daily_rollups` is no longer read for new-word counts (it counts surface forms). Its version state still gates the `vocabulary_visible` backfill.

## Vocabulary Word List

`/api/stats/vocabulary/words` (`query-vocabulary-words.ts`) folds `imm_words` into dictionary words on the server and filters, sorts, and pages there, because a client-side list can only hold a slice of the table. The fold is cached per connection until a write, the same way as the new-word aggregate. `titles` is computed for the returned rows; sorting or filtering by it needs every word's count (about 900ms), so that map tolerates five minutes of staleness. Names are detected with `isProperNounPos2`, since merged tokens store tags like `固有名詞|一般`.

## Rates

Rates are `null` when the underlying data is under 30 active minutes or 2,000 words, and charts draw that as a gap. Daily views rate a trailing 7 calendar days, volume-weighted, which is why the query loads days before the range. Week and month views rate each bucket.

## Title Filter

`title` narrows totals, charts, and patterns. `librarySummary` and `titleTimeline` stay complete so the page can switch titles. New words have no per-title record, so `learning` and `summary.current.newWords` are `null` while filtered.

## Metric Semantics

- subtitle-count stats now use Yomitan merged-token counts as the source of truth
- `tokensSeen` is the only active subtitle-count metric in tracker/session/rollup/query paths
- no whitespace/CJK-character fallback remains in the live stats path

## Contract

Stats data crosses the process boundary over HTTP only. Shared JSON models live in `src/types/stats-wire.ts`; endpoint request/response mappings and the client interface live in `src/types/stats-http-contract.ts`. The server and stats client both type-check against those files. `stats/src/types/stats.ts` is a compatibility re-export, not an independent contract copy.

The stats preload bridge remains limited to native window behavior such as confirmation-dialog layering. Do not add stats data request channels back to Electron IPC.

The stats UI should treat the trends payload as chart-ready data. Presentation-only work in the client is fine (including the title timeline layout), but rebuilding the main trend datasets from raw sessions should stay out of the render path.

For session detail timelines, omitting `limit` now means "return the full retained session telemetry/history". Explicit `limit` remains available for bounded callers, but the default stats UI path should not trim long sessions to the newest 200 samples.
