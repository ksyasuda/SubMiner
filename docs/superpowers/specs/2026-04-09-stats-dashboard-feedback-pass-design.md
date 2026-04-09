# Stats Dashboard Feedback Pass — Design

Date: 2026-04-09
Scope: Stats dashboard UX follow-ups from user feedback (items 1–7).
Delivery: **Single PR**, broken into logically scoped commits.

## Goals

Address seven concrete pieces of feedback against the Statistics menu:

1. Library — collapse episodes behind a per-series dropdown.
2. Sessions — roll up multiple sessions of the same episode within a day.
3. Trends — add a 365d range option.
4. Library — delete an episode (video) from its detail view.
5. Vocabulary — tighten spacing between word and reading in the Top 50 table.
6. Episode detail — hide cards whose Anki notes have been deleted.
7. Trend/watch charts — add gridlines, fix tick legibility, unify theming.

Out of scope for this pass: English-token ingestion cleanup and Overview stat-card drill-downs (feedback items 8 and 9). Those require a larger design decision and a migration respectively.

## Files touched (inventory)

Dashboard (`stats/src/`):
- `components/library/LibraryTab.tsx` — collapsible groups (item 1).
- `components/library/MediaDetailView.tsx`, `components/library/MediaHeader.tsx` — delete-episode action (item 4).
- `components/sessions/SessionsTab.tsx`, `components/library/MediaSessionList.tsx` — episode rollup (item 2).
- `components/trends/DateRangeSelector.tsx`, `hooks/useTrends.ts`, `lib/api-client.ts`, `lib/api-client.test.ts` — 365d (item 3).
- `components/vocabulary/FrequencyRankTable.tsx` — word/reading column collapse (item 5).
- `components/anime/EpisodeDetail.tsx` — filter deleted Anki cards (item 6).
- `components/trends/TrendChart.tsx`, `components/trends/StackedTrendChart.tsx`, `components/overview/WatchTimeChart.tsx`, `lib/chart-theme.ts` — chart clarity (item 7).
- New file: `stats/src/lib/session-grouping.ts` + `session-grouping.test.ts`.

Backend (`src/core/services/`):
- `immersion-tracker/query-trends.ts` — extend `TrendRange` and `TREND_DAY_LIMITS` (item 3).
- `immersion-tracker/__tests__/query.test.ts` — 365d coverage (item 3).
- `stats-server.ts` — passthrough if range validation lives here (check before editing).
- `__tests__/stats-server.test.ts` — 365d coverage (item 3).

## Commit plan

One PR, one feature per commit. Order picks low-risk mechanical changes first so failures in later commits don't block merging of earlier ones.

1. `feat(stats): add 365d range to trends dashboard` (item 3)
2. `fix(stats): tighten word/reading column in Top 50 table` (item 5)
3. `fix(stats): hide cards deleted from Anki in episode detail` (item 6)
4. `feat(stats): delete episode from library detail view` (item 4)
5. `feat(stats): collapsible series groups in library` (item 1)
6. `feat(stats): roll up same-episode sessions within a day` (item 2)
7. `feat(stats): gridlines and unified theme for trend charts` (item 7)

Each commit must pass `bun run typecheck`, `bun run test:fast`, and any change-specific checks listed below.

---

## Item 1 — Library collapsible series groups

### Current behavior

`LibraryTab.tsx` groups media via `groupMediaLibraryItems` and always renders the full grid of `MediaCard`s beneath each group header.

### Target behavior

Each group header becomes clickable. Groups with `items.length > 1` default to **collapsed**; single-video groups stay expanded (collapsing them would be visual noise).

### Implementation

- State: `const [collapsedGroups, setCollapsedGroups] = useState<Set<string>>(...)`. Initialize from `grouped` where `items.length > 1`.
- Toggle helper: `toggleGroup(key: string)` adds/removes from the set.
- Group header: wrap in a `<button>` with `aria-expanded` and a chevron icon (`▶`/`▼`). Keep the existing cover + title + subtitle layout inside the button.
- Children grid is conditionally rendered on `!collapsedGroups.has(group.key)`.
- Header summary (`N videos · duration · cards`) stays visible in both states so collapsed groups remain informative.

### Tests

- New `LibraryTab.test.tsx` (if not already present — check first) covering:
  - Multi-video group renders collapsed on first mount.
  - Single-video group renders expanded on first mount.
  - Clicking the header toggles visibility.
  - Header summary is visible in both states.

---

## Item 2 — Sessions episode rollup within a day

### Current behavior

`SessionsTab.tsx:10-24` groups sessions by day label only (`formatSessionDayLabel(startedAtMs)`). Multiple sessions of the same episode on the same day show as independent rows. `MediaSessionList.tsx` has the same problem inside the library detail view.

### Target behavior

Within each day, sessions with the same `videoId` collapse into one parent row showing combined totals. A chevron reveals the individual sessions. Single-session buckets render flat (no pointless nesting).

### Implementation

- New helper in `stats/src/lib/session-grouping.ts`:
  ```ts
  export interface SessionBucket {
    key: string;              // videoId as string, or `s-${sessionId}` for singletons
    videoId: number | null;
    sessions: SessionSummary[];
    totalActiveMs: number;
    totalCardsMined: number;
    representativeSession: SessionSummary; // most recent, for header display
  }
  export function groupSessionsByVideo(sessions: SessionSummary[]): SessionBucket[];
  ```
  Sessions missing a `videoId` become singleton buckets.

- `SessionsTab.tsx`: after day grouping, pipe each `daySessions` through `groupSessionsByVideo`. Render each bucket:
  - `sessions.length === 1`: existing `SessionRow` behavior, unchanged.
  - `sessions.length >= 2`: render a **bucket row** that looks like `SessionRow` but shows combined totals and session count (e.g. `3 sessions · 1h 24m · 12 cards`). Chevron state stored in a second `Set<string>` on bucket key. Expanded buckets render the child `SessionRow`s indented (`pl-8`) beneath the header.
- `MediaSessionList.tsx`: within the media detail view, a single video's sessions are all the same `videoId` by definition — grouping here is by day only, and within a day multiple sessions render nested under a day header. Re-use the same visual pattern; factor the bucket row into a shared `SessionBucketRow` component.

### Delete semantics

- Deleting a bucket header offers "Delete all N sessions in this group" (reuse `confirmDayGroupDelete` pattern with a bucket-specific message, or add `confirmBucketDelete`).
- Deleting an individual session from inside an expanded bucket keeps the existing single-delete flow.

### Tests

- `session-grouping.test.ts`:
  - Empty input → empty output.
  - All unique videos → N singleton buckets.
  - Two sessions same videoId → one bucket with correct totals and representative (most recent start time).
  - Missing videoId → singleton bucket keyed by sessionId.
- `SessionsTab.test.tsx` (extend or add) verifying the rendered bucket rows expand/collapse and delete hooks fire with the right ID set.

---

## Item 3 — 365d trends range

### Backend

`src/core/services/immersion-tracker/query-trends.ts`:
- `type TrendRange = '7d' | '30d' | '90d' | '365d' | 'all';`
- Add `'365d': 365` to `TREND_DAY_LIMITS`.
- `getTrendDayLimit` picks up the new key automatically because of the `Exclude<TrendRange, 'all'>` generic.

`src/core/services/stats-server.ts`:
- Search for any hardcoded range validation (e.g. allow-list in the trends route handler) and extend it.

### Frontend

- `hooks/useTrends.ts`: widen the `TimeRange` union.
- `components/trends/DateRangeSelector.tsx`: add `'365d'` to the options list. Display label stays as `365d`.
- `lib/api-client.ts` / `api-client.test.ts`: if the client validates ranges, add `365d`.

### Tests

- `query.test.ts`: extend the existing range table to cover `365d` returning 365 days of data.
- `stats-server.test.ts`: ensure the route accepts `range=365d`.
- `api-client.test.ts`: ensure the client emits the new range.

### Change-specific checks

- `bun run test:config` is not required here (no schema/defaults change).
- Run `bun run typecheck` + `bun run test:fast`.

---

## Item 4 — Delete episode from library detail

### Current behavior

`MediaDetailView.tsx` provides session-level delete only. The backend `deleteVideo` exists (`query-maintenance.ts:509`), the API is exposed at `stats-server.ts:559`, and `api-client.deleteVideo` is already wired (`stats/src/lib/api-client.ts:146`). `EpisodeList.tsx:46` already uses it from the anime tab.

### Target behavior

A "Delete Episode" action in `MediaHeader` (top-right, small, `text-ctp-red`), gated by `confirmEpisodeDelete(title)`. On success, call `onBack()` and make sure the parent `LibraryTab` refetches.

### Implementation

- Add an `onDeleteEpisode?: () => void` prop to `MediaHeader` and render the button only if provided.
- In `MediaDetailView`:
  - New handler `handleDeleteEpisode` that calls `apiClient.deleteVideo(videoId)`, then `onBack()`.
  - Reuse `confirmEpisodeDelete` from `stats/src/lib/delete-confirm.ts`.
- In `LibraryTab`:
  - `useMediaLibrary` returns fresh data on mount. The simplest fix: pass a `refresh` function from the hook (extend the hook if it doesn't already expose one) and call it when the detail view signals back.
  - Alternative: force a remount by incrementing a `libraryVersion` key on the library list. Prefer `refresh` for clarity.

### Tests

- Extend the existing `MediaDetailView.test.tsx`: mock `apiClient.deleteVideo`, click the new button, confirm `onBack` fires after success.
- `useMediaLibrary.test.ts`: if we add a `refresh` method, cover it.

---

## Item 5 — Vocabulary word/reading column collapse

### Current behavior

`FrequencyRankTable.tsx:110-144` uses a 5-column table: `Rank | Word | Reading | POS | Seen`. Word and Reading are auto-sized, producing a large gap.

### Target behavior

Merge Word + Reading into a single column titled "Word". Reading sits immediately after the headword in a muted, smaller style.

### Implementation

- Drop the `<th>Reading</th>` header and cell.
- Word cell becomes:
  ```tsx
  <td className="py-1.5 pr-3">
    <span className="text-ctp-text font-medium">{w.headword}</span>
    {reading && (
      <span className="text-ctp-subtext0 text-xs ml-1.5">
        【{reading}】
      </span>
    )}
  </td>
  ```
  where `reading = fullReading(w.headword, w.reading)` and differs from `headword`.
- Keep `fullReading` import from `reading-utils`.

### Tests

- Extend `FrequencyRankTable.test.tsx` (if present — otherwise add a focused test) to assert:
  - Headword renders.
  - Reading renders when different from headword.
  - Reading does not render when equal to headword.

---

## Item 6 — Hide Anki-deleted cards in Cards Mined

### Current behavior

`EpisodeDetail.tsx:109-147` iterates `cardEvents`, fetches note info via `ankiNotesInfo(allNoteIds)`, and for each `noteId` renders a row even if no matching `info` came back — the user sees an empty word with an "Open in Anki" button that leads nowhere.

### Target behavior

After `ankiNotesInfo` resolves:
- Drop `noteId`s that are not in the resolved map.
- Drop `cardEvents` whose `noteIds` list was non-empty but is now empty after filtering.
- Card events with a positive `cardsDelta` but no `noteIds` (legacy rollup path) still render as `+N cards` — we have no way to cross-reference them, so leave them alone.

### Implementation

- Compute `filteredCardEvents` as a `useMemo` depending on `data.cardEvents` and `noteInfos`.
- Iterate `filteredCardEvents` instead of `cardEvents` in the render.
- Surface a subtle note (optional, muted) "N cards hidden (deleted from Anki)" at the end of the list if any were filtered — helps the user understand why counts here diverge from session totals. Final decision on the note can be made at PR review; default: **show it**.

### Tests

- Add a test in `EpisodeDetail.test.tsx` (add the file if not present) that stubs `ankiNotesInfo` to return only a subset of notes and verifies the missing ones are not rendered.

### Other call sites

- Grep so far shows `ankiNotesInfo` is only used in `EpisodeDetail.tsx`. Re-verify before landing the commit; if another call site appears, apply the same filter.

---

## Item 7 — Trend/watch chart clarity pass

### Current behavior

`TrendChart.tsx`, `StackedTrendChart.tsx`, and `WatchTimeChart.tsx` render Recharts components with:
- No `CartesianGrid` → no horizontal reference lines.
- 9px axis ticks → borderline unreadable.
- Height 120 → cramped.
- Tooltip uses raw labels (`04/04` etc.).
- No shared theme object; each chart redefines colors and tooltip styles inline.

`stats/src/lib/chart-theme.ts` already exists and currently exports a single `CHART_THEME` constant with tick/tooltip colors and `barFill`. It will be extended, not replaced, to preserve existing consumers.

### Target behavior

All three charts share a theme, have horizontal gridlines, readable ticks, and sensible tooltips.

### Implementation

Extend `stats/src/lib/chart-theme.ts` with the additional shared defaults (keeping the existing `CHART_THEME` export intact so current consumers don't break):
```ts
export const CHART_THEME = {
  tick: '#a5adcb',
  tooltipBg: '#363a4f',
  tooltipBorder: '#494d64',
  tooltipText: '#cad3f5',
  tooltipLabel: '#b8c0e0',
  barFill: '#8aadf4',
  grid: '#494d64',
  axisLine: '#494d64',
} as const;

export const CHART_DEFAULTS = {
  height: 160,
  tickFontSize: 11,
  margin: { top: 8, right: 8, bottom: 0, left: 0 },
  grid: { strokeDasharray: '3 3', vertical: false },
} as const;

export const TOOLTIP_CONTENT_STYLE = {
  background: CHART_THEME.tooltipBg,
  border: `1px solid ${CHART_THEME.tooltipBorder}`,
  borderRadius: 6,
  color: CHART_THEME.tooltipText,
  fontSize: 12,
};
```

Apply to each chart:
- Import `CartesianGrid` from recharts.
- Insert `<CartesianGrid stroke={CHART_THEME.grid} {...CHART_DEFAULTS.grid} />` inside each chart container.
- `<XAxis tick={{ fontSize: CHART_DEFAULTS.tickFontSize, fill: CHART_THEME.tick }} />` and equivalent `YAxis`.
- `YAxis` gains `axisLine={{ stroke: CHART_THEME.axisLine }}`.
- `ResponsiveContainer` height changes from 120 → `CHART_DEFAULTS.height`.
- `Tooltip` `contentStyle` uses `TOOLTIP_CONTENT_STYLE`, and charts pass a `labelFormatter` when the label is a date key (e.g. show `Fri Apr 4`).

### Unit formatters

- `TrendChart` already accepts a `formatter` prop — extend usage sites to pass unit-aware formatters where they aren't already (`formatDuration`, `formatNumber`, etc.).

### Tests

- `chart-theme.test.ts` (if present — otherwise add a trivial snapshot to keep the shape stable).
- `TrendChart` snapshot/render tests: no regression, gridline element present.

---

## Verification gate

Before requesting code review, run:

```
bun run typecheck
bun run test:fast
bun run test:env
bun run test:runtime:compat   # dist-sensitive check for the charts
bun run build
bun run test:smoke:dist
```

No docs-site changes are planned in this spec; if `docs-site/` ends up touched (e.g. screenshots), also run `bun run docs:test` and `bun run docs:build`.

No config schema changes → `bun run test:config` and `bun run generate:config-example` are not required.

## Risks and open questions

- **MediaDetailView refresh**: `useMediaLibrary` may not expose a `refresh` function. If it doesn't, the simplest path is adding one; the alternative (keying a remount) works but is harder to test. Decide during implementation.
- **Session bucket delete UX**: "Delete all N sessions in this group" is powerful. The copy must make it clear the underlying sessions are being removed, not just the grouping. Reuse `confirmBucketDelete` wording from existing confirm helpers if possible.
- **Anki-deleted-cards hidden notice**: Showing a subtle "N cards hidden" footer is a call that can be made at PR review.
- **Bucket delete helper**: `confirmBucketDelete` does not currently exist in `delete-confirm.ts`. Implementation either adds it or reuses `confirmDayGroupDelete` with bucket-specific wording — decide during the session-rollup commit.

## Changelog entry

User-visible PR → needs a fragment under `changes/*.md`. Suggested title:
`Stats dashboard: collapsible series, session rollups, 365d trends, chart polish, episode delete.`
