# Library Summary Replaces Per-Day Trends — Design

**Status:** Draft
**Date:** 2026-04-09
**Scope:** `stats/` frontend, `src/core/services/immersion-tracker/query-trends.ts` backend

## Problem

The "Library — Per Day" section on the stats Trends tab (`stats/src/components/trends/TrendsTab.tsx:224-254`) renders six stacked-area charts — Videos, Watch Time, Cards, Words, Lookups, and Lookups/100w, each broken down per title per day.

In practice these charts are not useful:

- Most titles only have activity on one or two days in a window, so they render as isolated bumps on a noisy baseline.
- Stacking 7+ titles with mostly-zero days makes individual lines hard to follow.
- The top "Activity" and "Period Trends" sections already answer "what am I doing per day" globally.
- The "Library — Cumulative" section directly below already answers "which titles am I progressing through" with less noise.

The per-day section occupies significant vertical space without carrying its weight, and the user has confirmed it should be replaced.

## Goal

Replace the six per-day stacked charts with a single "Library — Summary" section that surfaces per-title aggregate statistics over the selected date range. The new view should make it trivially easy to answer: "For the selected window, which titles am I spending time on, how much mining output have they produced, and how efficient is my lookup rate on each?"

## Non-goals

- Changing the "Library — Cumulative" section (stays as-is).
- Changing the "Activity", "Period Trends", or "Patterns" sections.
- Adding a new API endpoint — the existing dashboard endpoint is extended in place.
- Renaming internal `anime*` data-model identifiers (`animeId`, `imm_anime`, etc.). Those stay per the convention established in `c5e778d7`; only new fields/types/user-visible strings use generic "title"/"library" wording.
- Supporting a true all-time library view on the Trends tab. If that's ever wanted, it belongs on a different tab.

## Solution Overview

Delete the "Library — Per Day" section. In its place, add "Library — Summary", composed of:

1. A horizontal-bar leaderboard chart of watch time per title (top 10, descending).
2. A sortable table of every title with activity in the selected window, with columns: Title, Watch Time, Videos, Sessions, Cards, Words, Lookups, Lookups/100w, Date Range.

Both controls are scoped to the top-of-page date range selector. The existing shared Anime Visibility filter continues to work — it now gates Summary + Cumulative instead of Per-Day + Cumulative.

## Backend

### New type

Add to `stats/src/types/stats.ts` and the backend query module:

```ts
type LibrarySummaryRow = {
  title: string;              // display title — anime series, YouTube video title, etc.
  watchTimeMin: number;       // sum(total_active_min) across the window
  videos: number;             // distinct video_id count
  sessions: number;           // session count from imm_sessions
  cards: number;              // sum(total_cards)
  words: number;              // sum(total_tokens_seen)
  lookups: number;            // sum(lookup_count) from imm_sessions
  lookupsPerHundred: number | null; // lookups / words * 100, null when words == 0
  firstWatched: number;       // min(rollup_day) as epoch day, within the window
  lastWatched: number;        // max(rollup_day) as epoch day, within the window
};
```

### Query changes in `src/core/services/immersion-tracker/query-trends.ts`

- Add `librarySummary: LibrarySummaryRow[]` to `TrendsDashboardQueryResult`.
- Populate it from a single aggregating query over `imm_daily_rollups` joined to `imm_videos` → `imm_anime`, filtered by `rollup_day` within the selected window. Session count and lookup count come from `imm_sessions` aggregated by `video_id` and then grouped by the parent library entry. Use a single query (or at most two joined/unioned) — no N+1.
- `imm_anime` is the generic library-grouping table; anime series, YouTube videos, and yt-dlp imports all land there. The internal table name stays `imm_anime`; only the new field uses generic naming.
- Return rows pre-sorted by `watchTimeMin` descending so the leaderboard is zero-cost and the table default sort matches.
- Emit `lookupsPerHundred: null` when `words == 0`.

### Removed from API response

- `animePerDay.lookups`
- `animePerDay.lookupsPerHundred`
- `animePerDay.episodes` (if no other consumer — verify during implementation)
- `animePerDay.watchTime` (if no other consumer — verify during implementation)
- `animePerDay.cards` (if no other consumer — verify during implementation)
- `animePerDay.words` (if no other consumer — verify during implementation)

If all four "if no other consumer" fields can be dropped, remove the `animePerDay` object from the response entirely along with the helpers that produce it. Check every use site in `stats/` before deleting.

## Frontend

### New component: `stats/src/components/trends/LibrarySummarySection.tsx`

Owns the header, leaderboard chart, visibility-filtered data, and the table. Keeps `TrendsTab.tsx` from growing. Component props: `{ rows: LibrarySummaryRow[]; hiddenTitles: ReadonlySet<string>; windowStart: Date; windowEnd: Date }`.

Internal state: `useState<{ column: ColumnId; direction: 'asc' | 'desc' }>` for sort, defaulting to `{ column: 'watchTimeMin', direction: 'desc' }`.

### Layout

Replaces `TrendsTab.tsx:224-254`:

```
[SectionHeader: "Library — Summary"]
[AnimeVisibilityFilter — unchanged, shared with Cumulative below]
[Card, col-span-full: Leaderboard — horizontal bar chart, ~260px tall]
[Card, col-span-full: Sortable table, auto height up to ~480px with internal scroll]
```

Both cards use the existing chart/card wrapper styling.

### Leaderboard chart

- ECharts horizontal bar chart (matches the rest of the page).
- Top 10 titles by watch time. If fewer titles have activity, render what's there.
- Y-axis: title, truncated with ellipsis at container width; full title on hover via ECharts tooltip.
- X-axis: minutes.
- Single series color: `#8aadf4` (matching the existing Watch Time color).
- Chart order is fixed at watch-time desc regardless of table sort — the leaderboard's meaning is fixed.

### Table

- Plain HTML `<table>` with Tailwind classes. No new deps.
- Columns, in order:
  1. **Title** — left-aligned, sticky, truncated with ellipsis, full title on hover.
  2. **Watch Time** — formatted `Xh Ym` when ≥60 min, else `Xm`.
  3. **Videos** — integer.
  4. **Sessions** — integer.
  5. **Cards** — integer.
  6. **Words** — integer.
  7. **Lookups** — integer.
  8. **Lookups/100w** — one decimal place, `—` when null.
  9. **Date Range** — `Mon D → Mon D` using the title's `firstWatched` / `lastWatched` within the window.
- Click a column header to sort; click again to reverse. Visual arrow on the active column.
- Numeric columns right-aligned.
- Null `lookupsPerHundred` sorts as the lowest value in both directions (consistent with "no data").
- Row hover highlight; no row click action (read-only view).
- Empty state: "No library activity in the selected window."

### Visibility filter integration

Hiding a title via `AnimeVisibilityFilter` removes it from both the leaderboard and the table. The filter's set of available titles is built from the union of titles that appear in `librarySummary` and the existing `animeCumulative.*` arrays (matches current behavior in `buildAnimeVisibilityOptions`).

### `TrendsTab.tsx` changes

- Remove the `filteredEpisodesPerAnime`, `filteredWatchTimePerAnime`, `filteredCardsPerAnime`, `filteredWordsPerAnime`, `filteredLookupsPerAnime`, `filteredLookupsPerHundredPerAnime` locals.
- Remove the six `<StackedTrendChart>` calls in the "Library — Per Day" section.
- Remove the `<SectionHeader>Library — Per Day</SectionHeader>` and the `<AnimeVisibilityFilter>` from that position.
- Insert `<SectionHeader>Library — Summary</SectionHeader>` + `<AnimeVisibilityFilter>` + `<LibrarySummarySection>` in the same place.
- Update `buildAnimeVisibilityOptions` input to use `librarySummary` titles instead of the six dropped `animePerDay.*` arrays.

## Data flow

1. `useTrends(range, groupBy)` calls `/api/stats/trends/dashboard`.
2. Response now includes `librarySummary` (sorted by watch time desc).
3. `TrendsTab` holds the shared `hiddenAnime` set (unchanged).
4. `LibrarySummarySection` receives `librarySummary` + `hiddenAnime`, filters out hidden rows, renders the leaderboard from the top-10 slice of the filtered list, renders the table from the filtered list with local sort state applied.
5. Date-range selector changes trigger a new fetch; `groupBy` toggle does not affect the summary section (it's always window-total).

## Edge cases

- **No activity in window:** Section renders header + empty-state card. Leaderboard card hidden. Visibility filter hidden.
- **One title only:** Leaderboard renders a single bar; table renders one row. No special-casing.
- **Title with zero words but non-zero lookups:** `lookupsPerHundred` is `null`, rendered as `—`. Sort treats null as lowest.
- **Title with zero cards/lookups/words but non-zero watch time:** Normal zero rendering, still shown.
- **Very long titles:** Ellipsis in chart y-axis labels and table title column; full title in `title` attribute / ECharts tooltip.
- **Mixed sources (anime + YouTube):** No special case — both land in `imm_anime` and are grouped uniformly.

## Testing

### Backend (`query-trends.ts`)

New unit tests, following the existing pattern:

1. Empty window returns `librarySummary: []`.
2. Single title with a few rollups: all aggregates are correct; `firstWatched`/`lastWatched` match the bounding days within the window.
3. Multiple titles: rows returned sorted by watch time desc.
4. Mixed sources (anime-style + YouTube-style entries in `imm_anime`): both appear in the summary with their own aggregates.
5. Title with `words == 0`: `lookupsPerHundred` is `null`.
6. Date range excludes some rollups: excluded rollups are not counted; `firstWatched`/`lastWatched` reflect only within-window activity.
7. `sessions` and `lookups` come from `imm_sessions`, not `imm_daily_rollups`, and are correctly attributed to the parent library entry.

### Frontend

- Existing Trends tab smoke test should continue to pass after wiring.
- Optional: a targeted render test for `LibrarySummarySection` (empty state, single title, sort toggle, visibility filter interaction). Not required for merge if the smoke test exercises the happy path.

## Release / docs

- One fragment in `changes/*.md` summarizing the replacement.
- No user-facing docs (`docs-site/`) changes unless the per-day section was documented there — verify during implementation.

## Open items

None.
