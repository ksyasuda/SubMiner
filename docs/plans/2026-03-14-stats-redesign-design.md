# Stats Dashboard Redesign — Anime-Centric Approach

**Date**: 2026-03-14
**Status**: Approved

## Motivation

The current stats dashboard tracks metrics that aren't particularly useful (words seen as a hero stat, word clouds). The data model now supports anime-level tracking (`imm_anime`, `imm_videos` with `parsed_episode`), subtitle line storage (`imm_subtitle_lines`), and word/kanji occurrence mapping (`imm_word_line_occurrences`, `imm_kanji_line_occurrences`). The dashboard should be restructured around anime as the primary unit, with sessions, episodes, and rollups as the core metrics.

## Data Model (already in place)

- `imm_anime` — anime-level: title, AniList ID, romaji/english/native titles, metadata
- `imm_videos` — episode-level: `anime_id`, `parsed_episode`, `parsed_season`
- `imm_sessions` — session-level: linked to video
- `imm_subtitle_lines` — line-level: linked to session, video, anime
- `imm_word_line_occurrences` / `imm_kanji_line_occurrences` — word/kanji → line mapping
- `imm_media_art` — cover art + `episodes_total`
- `imm_daily_rollups` / `imm_monthly_rollups` — aggregated metrics
- `imm_words` — POS data: `part_of_speech`, `pos1`, `pos2`, `pos3`

## Tab Structure (5 tabs)

### 1. Overview

**Hero Stats** (6 cards):
- Watch time today
- Cards mined today
- Sessions today
- Episodes watched today
- Current streak (days)
- Active anime (titles with sessions in last 30 days)

**14-day Watch Time Chart**: Bar chart (keep existing).

**Streak Calendar**: GitHub-contributions-style heatmap, last 90 days, colored by watch time intensity.

**Tracking Snapshot** (secondary stats): Total sessions, total episodes, all-time hours, active days, total cards.

**Recent Activity Feed**: Last 10 sessions grouped by day — anime title + cover art thumbnail, episode number, duration, cards mined.

Removed from Overview: 14-day words chart, "words today", "words this week" hero stats.

### 2. Anime (replaces Library)

**Grid View**:
- Responsive card grid with cover art
- Each card: title, progress bar (episodes watched / `episodes_total`), watch time, cards mined
- Search/filter by title
- Sort: last watched, watch time, cards mined, progress %

**Anime Detail View** (click into card):
- Header: cover art, titles (romaji/english/native), AniList link if available
- Progress: episode progress bar + "X / Y episodes"
- Stats row: total watch time, cards mined, words seen, lookup hit rate, avg session length
- Episode list: table of episodes (from `imm_videos`), each showing episode number, session count, watch time, cards, last watched date
- Watch time chart: bar chart over time (14d/30d/90d toggle)
- Words from this anime: top words learned from this show (via `imm_word_line_occurrences` → `imm_subtitle_lines` → `anime_id`), clickable to vocab detail
- Mining efficiency: cards per hour / cards per episode trend

### 3. Trends

**Existing charts (keep all 9)**:
1. Watch Time (min) — bar
2. Tracked Cards — bar
3. Words Seen — bar
4. Sessions — line
5. Avg Session (min) — line
6. Cards per Hour — line
7. Lookup Hit Rate (%) — line
8. Rolling 7d Watch Time — line
9. Rolling 7d Cards — line

**New charts (6)**:
10. Episodes watched per day/week
11. Anime completion progress over time (cumulative episodes / total across all anime)
12. New anime started over time (first session per anime by date)
13. Watch time per anime (stacked bar — top 5 anime + "other")
14. Streak history (visual streak timeline — active vs gap periods)
15. Cards per episode trend

**Controls**: Time range selector (7d/30d/90d/all), group by (day/month).

### 4. Vocabulary

**Hero Stats** (4 cards):
- Unique words (excluding particles/noise via POS filter)
- Unique kanji
- New this week
- Avg frequency

**Filters/Controls**:
- POS filter toggle: hide particles, single-char tokens by default (toggleable)
- Sort: by frequency / last seen / first seen
- Search by word/reading

**Word List**: Grid/table of words — headword, reading, POS tag, frequency. Each word is clickable.

**Word Detail Panel** (slide-out or modal):
- Headword, reading, POS (part_of_speech, pos1, pos2, pos3)
- Frequency + first/last seen dates
- Anime appearances: which anime this word appeared in, frequency per anime
- Example lines: actual subtitle lines where the word was used
- Similar words: words sharing same kanji or reading

**Kanji Section**: Same pattern — clickable kanji grid, detail panel with frequency, anime appearances, example lines, words using this kanji.

**Charts**: Top repeated words bar chart, new words by day timeline.

### 5. Sessions

**Session List**: Chronological, grouped by day.
- Each row: anime title + episode, cover art thumbnail, duration (active/total), cards mined, lines seen, lookup rate
- Expandable detail: session timeline chart (words/cards over time), event log (pauses, seeks, lookups, cards mined)
- Filters: by anime title, date range

Based on existing hidden `SessionsTab` component with anime/episode context added.

## Backend Changes Needed

### New API Endpoints
- `GET /api/stats/anime` — list all anime with episode counts, watch time, progress
- `GET /api/stats/anime/:animeId` — anime detail: episodes, stats, recent sessions
- `GET /api/stats/anime/:animeId/words` — top words from this anime
- `GET /api/stats/vocabulary/:wordId` — word detail: POS, frequency, anime appearances, example lines, similar words
- `GET /api/stats/kanji/:kanjiId` — kanji detail: frequency, anime appearances, example lines, words using this kanji

### Modified API Endpoints
- `GET /api/stats/vocabulary` — add POS fields to response, support POS filtering query param
- `GET /api/stats/overview` — add episodes today, active anime count
- `GET /api/stats/daily-rollups` — add episode count data for new trend charts

### New Query Functions
- Anime-level aggregation: episodes per anime, watch time per anime, cards per anime
- Word/kanji occurrence lookups: join through `imm_word_line_occurrences` → `imm_subtitle_lines` → `imm_anime`
- Streak calendar data: daily activity map for last 90 days
- Episode-level trend data: episodes per day for trend charts
- Stacked watch time: per-anime daily breakdown
