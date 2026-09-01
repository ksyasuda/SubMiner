# Stats Trends Data Flow

read_when: touching stats trend charts, changing stats API payloads, or debugging dashboard performance

## Summary

Trend charts now consume one chart-oriented backend payload from `/api/stats/trends/dashboard`.

## Why

- remove repeated client-side dataset rebuilding in `TrendsTab`
- collapse multiple network round-trips into one request
- keep heavy chart shaping close to tracker/query logic

## Data Sources

- rollup-backed:
  - activity charts
  - cumulative watch/cards/tokens/sessions trends
  - per-anime watch/cards/tokens/episodes series
- session-metric-backed:
  - lookup trends
  - lookup rate trends
  - watch-time by day-of-week/hour
- vocabulary-backed:
  - new-words trend reads permanent daily lexical rollups
  - rollup rows count only vocabulary-visible tokens and normalize mixed legacy timestamp units
  - a persisted rollup version invalidates stale materializations and triggers an atomic background rebuild

## Metric Semantics

- subtitle-count stats now use Yomitan merged-token counts as the source of truth
- `tokensSeen` is the only active subtitle-count metric in tracker/session/rollup/query paths
- no whitespace/CJK-character fallback remains in the live stats path

## Contract

Stats data crosses the process boundary over HTTP only. Shared JSON models live in `src/types/stats-wire.ts`; endpoint request/response mappings and the client interface live in `src/types/stats-http-contract.ts`. The server and stats client both type-check against those files. `stats/src/types/stats.ts` is a compatibility re-export, not an independent contract copy.

The stats preload bridge remains limited to native window behavior such as confirmation-dialog layering. Do not add stats data request channels back to Electron IPC.

The stats UI should treat the trends payload as chart-ready data. Presentation-only work in the client is fine, but rebuilding the main trend datasets from raw sessions should stay out of the render path.

For session detail timelines, omitting `limit` now means "return the full retained session telemetry/history". Explicit `limit` remains available for bounded callers, but the default stats UI path should not trim long sessions to the newest 200 samples.
