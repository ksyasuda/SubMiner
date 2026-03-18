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
  - cumulative watch/cards/words/sessions trends
  - per-anime watch/cards/words/episodes series
- session-metric-backed:
  - lookup trends
  - lookup rate trends
  - watch-time by day-of-week/hour
- vocabulary-backed:
  - new-words trend

## Contract

The stats UI should treat the trends payload as chart-ready data. Presentation-only work in the client is fine, but rebuilding the main trend datasets from raw sessions should stay out of the render path.
