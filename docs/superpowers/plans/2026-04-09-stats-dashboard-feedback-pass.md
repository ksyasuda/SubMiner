# Stats Dashboard Feedback Pass Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Land seven UX/correctness improvements to the SubMiner stats dashboard in a single PR with one logical commit per task.

**Architecture:** All work lives in `stats/src/` (React + Vite + Tailwind UI) and `src/core/services/immersion-tracker/` (sqlite-backed query layer for the 365d range only). No schema changes, no migrations. Each task is independently testable; helpers get their own files with focused unit tests.

**Tech Stack:** Bun, TypeScript, React 19, Recharts, Tailwind, sqlite via `node:sqlite`, `bun:test`.

**Spec:** `docs/superpowers/specs/2026-04-09-stats-dashboard-feedback-pass-design.md`

---

## Working agreements

- One commit per task. Commit message conventions follow recent history (`feat(stats):` / `fix(stats):` / `docs:`).
- Run `bun run typecheck` after every commit; run `bun run typecheck:stats` after stats UI commits.
- Use Bun, not Node: `bun test <file>`, not `npx jest`.
- Ranges or files cited as `path:line` are valid as of branch `stats-update`. If something has moved, re-grep before editing.
- Tests for stats UI files live alongside their source (e.g. `LibraryTab.tsx` → `LibraryTab.test.tsx`). Use `bun test stats/src/path/to/file.test.tsx` to run a single file.
- Frequently committing means: each task is its own commit. Don't squash.

## Pre-flight (do this once before Task 1)

- [ ] **Step 1: Verify branch and clean tree**

  Run: `git status && git log --oneline -3`
  Expected: branch `stats-update`, working tree clean except for whatever you're about to do, last commit is the spec commit `82d58a57`.

- [ ] **Step 2: Verify baseline tests pass**

  Run: `bun run typecheck && bun run typecheck:stats`
  Expected: both succeed.

- [ ] **Step 3: Confirm bun test runs single stats files**

  Run: `bun test stats/src/lib/api-client.test.ts`
  Expected: tests pass.

---

## Task 1: 365d range — backend type extension

**Files:**
- Modify: `src/core/services/immersion-tracker/query-trends.ts:16` and `src/core/services/immersion-tracker/query-trends.ts:84-88`
- Test: `src/core/services/immersion-tracker/__tests__/query.test.ts`

- [ ] **Step 1: Read the existing range table test**

  Run: `bun test src/core/services/immersion-tracker/__tests__/query.test.ts -t '365d' 2>&1 | head -20`
  Expected: no test named `365d`. Locate existing range coverage by reading the file (`grep -n '7d\|30d\|90d' src/core/services/immersion-tracker/__tests__/query.test.ts`) so you can mirror its style.

- [ ] **Step 2: Add a failing test that asserts `365d` returns up to 365 day buckets**

  Edit `src/core/services/immersion-tracker/__tests__/query.test.ts`. Find an existing trend range test (search for `'90d'`) and add a new sibling test that:
  - Seeds 400 days of synthetic daily activity (or whatever the existing helpers use).
  - Calls the trends query with `range: '365d', groupBy: 'day'`.
  - Asserts the returned `watchTimeByDay.length === 365`.
  - Mirrors the assertions style of the existing 90d test exactly.

- [ ] **Step 3: Run the new test to verify it fails**

  Run: `bun test src/core/services/immersion-tracker/__tests__/query.test.ts -t '365d'`
  Expected: TypeScript compile error or runtime failure because `'365d'` is not assignable to `TrendRange`.

- [ ] **Step 4: Extend `TrendRange` and `TREND_DAY_LIMITS`**

  In `src/core/services/immersion-tracker/query-trends.ts`:
  - Line 16: change to `type TrendRange = '7d' | '30d' | '90d' | '365d' | 'all';`
  - Line 84-88: add `'365d': 365,` so the map becomes:
    ```ts
    const TREND_DAY_LIMITS: Record<Exclude<TrendRange, 'all'>, number> = {
      '7d': 7,
      '30d': 30,
      '90d': 90,
      '365d': 365,
    };
    ```

- [ ] **Step 5: Run the new test to verify it passes**

  Run: `bun test src/core/services/immersion-tracker/__tests__/query.test.ts -t '365d'`
  Expected: PASS.

- [ ] **Step 6: Run the full query test file to verify no regressions**

  Run: `bun test src/core/services/immersion-tracker/__tests__/query.test.ts`
  Expected: all tests pass.

- [ ] **Step 7: Commit**

  ```bash
  git add src/core/services/immersion-tracker/query-trends.ts \
          src/core/services/immersion-tracker/__tests__/query.test.ts
  git commit -m "feat(stats): support 365d range in trends query"
  ```

---

## Task 2: 365d range — server route allow-list

**Files:**
- Modify: `src/core/services/stats-server.ts` (search for trends route handler — look for `/api/stats/trends` or `getTrendsDashboard`)
- Test: `src/core/services/__tests__/stats-server.test.ts`

- [ ] **Step 1: Locate the trends route in `stats-server.ts`**

  Run: `grep -n 'trends\|TrendRange' src/core/services/stats-server.ts`
  Read the surrounding code. If the route delegates straight through to `tracker.getTrendsDashboard(range, groupBy)` without an allow-list, **this entire task is a no-op** — skip ahead to Task 3 and document in the commit message of Task 3 that no server changes were needed. If there *is* an allow-list (e.g. a `validRanges` array), continue.

- [ ] **Step 2: Add a failing test for `range=365d`**

  In `src/core/services/__tests__/stats-server.test.ts`, find the existing trends route test (search for `'/api/stats/trends'`). Add a sibling case that issues a request with `range=365d` and asserts the response is 200 (not 400).

- [ ] **Step 3: Run the test to verify it fails**

  Run: `bun test src/core/services/__tests__/stats-server.test.ts -t '365d'`
  Expected: FAIL because `365d` isn't in the allow-list.

- [ ] **Step 4: Extend the allow-list**

  Add `'365d'` to the `validRanges`/`allowedRanges` array (whatever it is named) so it sits next to `'90d'`.

- [ ] **Step 5: Re-run the test**

  Run: `bun test src/core/services/__tests__/stats-server.test.ts -t '365d'`
  Expected: PASS.

- [ ] **Step 6: Run the full server test file**

  Run: `bun test src/core/services/__tests__/stats-server.test.ts`
  Expected: all tests pass.

- [ ] **Step 7: Commit (only if step 1 found an allow-list)**

  ```bash
  git add src/core/services/stats-server.ts \
          src/core/services/__tests__/stats-server.test.ts
  git commit -m "feat(stats): allow 365d trends range in HTTP route"
  ```

---

## Task 3: 365d range — frontend client and selector

**Files:**
- Modify: `stats/src/lib/api-client.ts`
- Modify: `stats/src/lib/api-client.test.ts`
- Modify: `stats/src/hooks/useTrends.ts:5`
- Modify: `stats/src/components/trends/DateRangeSelector.tsx:56`

- [ ] **Step 1: Locate range usage in the api-client**

  Run: `grep -n 'TrendRange\|range\|7d\|90d' stats/src/lib/api-client.ts | head -20`
  Identify whether the client validates ranges or simply passes them through. Mirror your test/edit accordingly.

- [ ] **Step 2: Add a failing test in `api-client.test.ts`**

  Add a test case that calls `apiClient.getTrendsDashboard('365d', 'day')` (or whatever the public method is named), stubs `fetch`, and asserts the URL contains `range=365d`. Mirror the existing 90d test if there is one.

- [ ] **Step 3: Run the new test to verify it fails**

  Run: `bun test stats/src/lib/api-client.test.ts -t '365d'`
  Expected: FAIL on type-narrowing.

- [ ] **Step 4: Widen the client `TrendRange` union**

  In `stats/src/lib/api-client.ts`, find any `TrendRange`-shaped union and add `'365d'`. If the client re-imports the type from elsewhere, no edit needed beyond the consumer test.

- [ ] **Step 5: Update `useTrends.ts:5`**

  Change `export type TimeRange = '7d' | '30d' | '90d' | 'all';` to `export type TimeRange = '7d' | '30d' | '90d' | '365d' | 'all';`.

- [ ] **Step 6: Add `365d` to the `DateRangeSelector` segmented control**

  In `stats/src/components/trends/DateRangeSelector.tsx:56`, change:
  ```tsx
  options={['7d', '30d', '90d', 'all'] as TimeRange[]}
  ```
  to:
  ```tsx
  options={['7d', '30d', '90d', '365d', 'all'] as TimeRange[]}
  ```

- [ ] **Step 7: Run the new client test**

  Run: `bun test stats/src/lib/api-client.test.ts -t '365d'`
  Expected: PASS.

- [ ] **Step 8: Typecheck the stats UI**

  Run: `bun run typecheck:stats`
  Expected: succeeds.

- [ ] **Step 9: Commit**

  ```bash
  git add stats/src/lib/api-client.ts stats/src/lib/api-client.test.ts \
          stats/src/hooks/useTrends.ts stats/src/components/trends/DateRangeSelector.tsx
  git commit -m "feat(stats): expose 365d trends range in dashboard UI"
  ```

---

## Task 4: Vocabulary Top 50 — collapse word/reading column

**Files:**
- Modify: `stats/src/components/vocabulary/FrequencyRankTable.tsx:110-144`
- Test: create `stats/src/components/vocabulary/FrequencyRankTable.test.tsx` if not present (check first with `ls stats/src/components/vocabulary/`)

- [ ] **Step 1: Check whether a test file exists**

  Run: `ls stats/src/components/vocabulary/FrequencyRankTable.test.tsx 2>/dev/null || echo "missing"`
  If missing, you'll create it in step 2.

- [ ] **Step 2: Write the failing test**

  Create or extend `stats/src/components/vocabulary/FrequencyRankTable.test.tsx` with:
  ```tsx
  import { render, screen } from '@testing-library/react';
  import { describe, it, expect } from 'bun:test';
  import { FrequencyRankTable } from './FrequencyRankTable';
  import type { VocabularyEntry } from '../../types/stats';

  function makeEntry(over: Partial<VocabularyEntry>): VocabularyEntry {
    return {
      wordId: 1,
      headword: '日本語',
      word: '日本語',
      reading: 'にほんご',
      frequency: 5,
      frequencyRank: 100,
      animeCount: 1,
      partOfSpeech: null,
      firstSeen: 0,
      lastSeen: 0,
      ...over,
    } as VocabularyEntry;
  }

  describe('FrequencyRankTable', () => {
    it('renders headword and reading inline in a single column (no separate Reading header)', () => {
      const entry = makeEntry({});
      render(<FrequencyRankTable words={[entry]} knownWords={new Set()} />);
      // Reading should be visually associated with the headword, not in its own column.
      expect(screen.queryByRole('columnheader', { name: 'Reading' })).toBeNull();
      expect(screen.getByText('日本語')).toBeTruthy();
      expect(screen.getByText(/にほんご/)).toBeTruthy();
    });

    it('omits reading when reading equals headword', () => {
      const entry = makeEntry({ headword: 'カレー', word: 'カレー', reading: 'カレー' });
      render(<FrequencyRankTable words={[entry]} knownWords={new Set()} />);
      // Headword still renders; no bracketed reading line for the duplicate.
      expect(screen.getByText('カレー')).toBeTruthy();
      expect(screen.queryByText(/【カレー】/)).toBeNull();
    });
  });
  ```

  Note: this assumes `@testing-library/react` is already a dev dep — confirm with `grep '@testing-library/react' stats/package.json /Users/sudacode/projects/japanese/SubMiner/package.json`. If it's not in `stats/`, run the test with `bun test` from repo root since tooling may resolve from the parent. If the project's existing component tests use a different render helper (check `MediaDetailView.test.tsx` for the pattern), copy that pattern instead.

- [ ] **Step 3: Run the new test to verify it fails**

  Run: `bun test stats/src/components/vocabulary/FrequencyRankTable.test.tsx`
  Expected: FAIL — the current component still renders a Reading `<th>`.

- [ ] **Step 4: Modify `FrequencyRankTable.tsx`**

  Replace the `<th>Reading</th>` header column and the corresponding `<td>` in the body. The new shape:

  Header (around line 113-119):
  ```tsx
  <thead>
    <tr className="text-xs text-ctp-overlay2 border-b border-ctp-surface1">
      <th className="text-left py-2 pr-3 font-medium w-16">Rank</th>
      <th className="text-left py-2 pr-3 font-medium">Word</th>
      <th className="text-left py-2 pr-3 font-medium w-20">POS</th>
      <th className="text-right py-2 font-medium w-20">Seen</th>
    </tr>
  </thead>
  ```

  Body row (around line 122-141):
  ```tsx
  <tr
    key={w.wordId}
    onClick={() => onSelectWord?.(w)}
    className="border-b border-ctp-surface1 last:border-0 cursor-pointer hover:bg-ctp-surface1/50 transition-colors"
  >
    <td className="py-1.5 pr-3 font-mono tabular-nums text-ctp-peach text-xs">
      #{w.frequencyRank!.toLocaleString()}
    </td>
    <td className="py-1.5 pr-3">
      <span className="text-ctp-text font-medium">{w.headword}</span>
      {(() => {
        const reading = fullReading(w.headword, w.reading);
        if (!reading || reading === w.headword) return null;
        return (
          <span className="text-ctp-subtext0 text-xs ml-1.5">
            【{reading}】
          </span>
        );
      })()}
    </td>
    <td className="py-1.5 pr-3">
      {w.partOfSpeech && <PosBadge pos={w.partOfSpeech} />}
    </td>
    <td className="py-1.5 text-right font-mono tabular-nums text-ctp-blue text-xs">
      {w.frequency}x
    </td>
  </tr>
  ```

- [ ] **Step 5: Run the test to verify it passes**

  Run: `bun test stats/src/components/vocabulary/FrequencyRankTable.test.tsx`
  Expected: PASS.

- [ ] **Step 6: Typecheck**

  Run: `bun run typecheck:stats`
  Expected: succeeds.

- [ ] **Step 7: Commit**

  ```bash
  git add stats/src/components/vocabulary/FrequencyRankTable.tsx \
          stats/src/components/vocabulary/FrequencyRankTable.test.tsx
  git commit -m "fix(stats): collapse word and reading into one column in Top 50 table"
  ```

---

## Task 5: Episode detail — filter Anki-deleted cards

**Files:**
- Modify: `stats/src/components/anime/EpisodeDetail.tsx:109-147`
- Test: create `stats/src/components/anime/EpisodeDetail.test.tsx` if not present

- [ ] **Step 1: Confirm `ankiNotesInfo` is only consumed in `EpisodeDetail.tsx`**

  Run: `grep -rn 'ankiNotesInfo' stats/src`
  Expected: only `EpisodeDetail.tsx`. If anything else turns up, this task must also patch that consumer.

- [ ] **Step 2: Write the failing test**

  Create `stats/src/components/anime/EpisodeDetail.test.tsx` (copy the import/setup pattern from the closest existing component test like `MediaDetailView.test.tsx`):

  ```tsx
  import { render, screen, waitFor } from '@testing-library/react';
  import { describe, it, expect, mock, beforeEach } from 'bun:test';
  import { EpisodeDetail } from './EpisodeDetail';

  // Mock the stats client. Mirror the mocking style used in MediaDetailView.test.tsx.
  // The key behavior: ankiNotesInfo only returns one of the two requested noteIds.

  describe('EpisodeDetail card filtering', () => {
    beforeEach(() => {
      // reset mocks
    });

    it('hides card events whose Anki notes have been deleted', async () => {
      // Stub getStatsClient().getEpisodeDetail to return two cardEvents,
      // each with one noteId.
      // Stub ankiNotesInfo to return only the first noteId.
      // Render <EpisodeDetail videoId={1} />.
      // Wait for the cards to load.
      // Assert exactly one card row is visible.
      // Assert the surviving event's expression renders.
    });
  });
  ```

  Implementation note: the existing `EpisodeDetail.tsx` calls `getStatsClient()` directly. Look at how `MediaDetailView.test.tsx` handles this — there's already an established mocking pattern. Copy it. If it uses `mock.module('../../hooks/useStatsApi', ...)`, do the same.

- [ ] **Step 3: Run the test to verify it fails**

  Run: `bun test stats/src/components/anime/EpisodeDetail.test.tsx`
  Expected: FAIL — both card rows currently render even when one note is missing.

- [ ] **Step 4: Add the filter to `EpisodeDetail.tsx`**

  At the top of the render section (after `const { sessions, cardEvents } = data;` around line 73), insert:

  ```tsx
  const filteredCardEvents = cardEvents
    .map((ev) => {
      if (ev.noteIds.length === 0) {
        // Legacy rollup events with no noteIds — leave alone.
        return ev;
      }
      const survivingNoteIds = ev.noteIds.filter((id) => noteInfos.has(id));
      return { ...ev, noteIds: survivingNoteIds };
    })
    .filter((ev) => {
      // Drop events that originally had noteIds but lost them all.
      return ev.noteIds.length > 0 || ev.cardsDelta > 0;
    });

  // Track how many were hidden so we can surface a small footer.
  const hiddenCardCount = cardEvents.reduce((acc, ev) => {
    if (ev.noteIds.length === 0) return acc;
    const dropped = ev.noteIds.filter((id) => !noteInfos.has(id)).length;
    return acc + dropped;
  }, 0);
  ```

  Then change the JSX iteration from `cardEvents.map(...)` to `filteredCardEvents.map(...)` (one occurrence around line 113), and after the `</div>` closing the cards-mined section, add:

  ```tsx
  {hiddenCardCount > 0 && (
    <div className="px-3 pb-3 -mt-1 text-[10px] text-ctp-overlay2 italic">
      {hiddenCardCount} card{hiddenCardCount === 1 ? '' : 's'} hidden (deleted from Anki)
    </div>
  )}
  ```

  Place that footer immediately before the closing `</div>` of the bordered cards-mined section, so it stays scoped to that block.

  **Important:** the filter only fires once `noteInfos` has been populated. While `noteInfos` is still empty (initial load before the second fetch resolves), every card with noteIds would be filtered out — that's wrong. Guard the filter so that it only runs after the noteInfos fetch has completed. The simplest signal: track `noteInfosLoaded: boolean` next to `noteInfos`, set it `true` in the `.then` callback, and only apply filtering when `noteInfosLoaded || allNoteIds.length === 0`.

  Concrete change near line 22:
  ```tsx
  const [noteInfos, setNoteInfos] = useState<Map<number, NoteInfo>>(new Map());
  const [noteInfosLoaded, setNoteInfosLoaded] = useState(false);
  ```

  Inside the existing `useEffect` (around line 36-46), set the loaded flag:
  ```tsx
  if (allNoteIds.length > 0) {
    getStatsClient()
      .ankiNotesInfo(allNoteIds)
      .then((notes) => {
        if (cancelled) return;
        const map = new Map<number, NoteInfo>();
        for (const note of notes) {
          const expr = note.preview?.word ?? '';
          map.set(note.noteId, { noteId: note.noteId, expression: expr });
        }
        setNoteInfos(map);
        setNoteInfosLoaded(true);
      })
      .catch((err) => {
        console.warn('Failed to fetch Anki note info:', err);
        if (!cancelled) setNoteInfosLoaded(true); // unblock so we don't hide everything
      });
  } else {
    setNoteInfosLoaded(true);
  }
  ```

  And gate the filter:
  ```tsx
  const filteredCardEvents = noteInfosLoaded
    ? cardEvents
        .map((ev) => {
          if (ev.noteIds.length === 0) return ev;
          const survivingNoteIds = ev.noteIds.filter((id) => noteInfos.has(id));
          return { ...ev, noteIds: survivingNoteIds };
        })
        .filter((ev) => ev.noteIds.length > 0 || ev.cardsDelta > 0)
    : cardEvents;
  ```

- [ ] **Step 5: Run the test to verify it passes**

  Run: `bun test stats/src/components/anime/EpisodeDetail.test.tsx`
  Expected: PASS.

- [ ] **Step 6: Add a second test for the loading-state guard**

  Extend the test to assert that, before `ankiNotesInfo` resolves, both card rows still appear (so we don't briefly flash an empty list). Then verify that after resolution, the deleted one disappears.

- [ ] **Step 7: Run both tests**

  Run: `bun test stats/src/components/anime/EpisodeDetail.test.tsx`
  Expected: PASS.

- [ ] **Step 8: Typecheck**

  Run: `bun run typecheck:stats`
  Expected: succeeds.

- [ ] **Step 9: Commit**

  ```bash
  git add stats/src/components/anime/EpisodeDetail.tsx \
          stats/src/components/anime/EpisodeDetail.test.tsx
  git commit -m "fix(stats): hide cards deleted from Anki in episode detail"
  ```

---

## Task 6: Library detail — delete episode action

**Files:**
- Modify: `stats/src/components/library/MediaHeader.tsx`
- Modify: `stats/src/components/library/MediaDetailView.tsx`
- Modify: `stats/src/hooks/useMediaLibrary.ts`
- Modify: `stats/src/components/library/LibraryTab.tsx`
- Test: extend `stats/src/components/library/MediaDetailView.test.tsx`
- Test: extend or create `stats/src/hooks/useMediaLibrary.test.ts`

- [ ] **Step 1: Add a failing test for the delete button in `MediaDetailView.test.tsx`**

  Read `stats/src/components/library/MediaDetailView.test.tsx` first to see the test scaffolding. Then add a new test:

  ```tsx
  it('deletes the episode and calls onBack when the delete button is clicked', async () => {
    const onBack = mock(() => {});
    const deleteVideo = mock(async () => {});
    // Stub apiClient.deleteVideo with the mock above (mirror existing stub patterns).
    // Stub useMediaDetail to return a populated detail object.
    // Stub window.confirm to return true.
    render(<MediaDetailView videoId={42} onBack={onBack} />);
    // Wait for the header to render.
    const button = await screen.findByRole('button', { name: /delete episode/i });
    button.click();
    await waitFor(() => expect(deleteVideo).toHaveBeenCalledWith(42));
    await waitFor(() => expect(onBack).toHaveBeenCalled());
  });
  ```

- [ ] **Step 2: Run the failing test**

  Run: `bun test stats/src/components/library/MediaDetailView.test.tsx -t 'delete'`
  Expected: FAIL because no delete button exists.

- [ ] **Step 3: Add `onDeleteEpisode` prop to `MediaHeader`**

  In `stats/src/components/library/MediaHeader.tsx`:

  ```tsx
  interface MediaHeaderProps {
    detail: NonNullable<MediaDetailData['detail']>;
    initialKnownWordsSummary?: {
      totalUniqueWords: number;
      knownWordCount: number;
    } | null;
    onDeleteEpisode?: () => void;
  }

  export function MediaHeader({
    detail,
    initialKnownWordsSummary = null,
    onDeleteEpisode,
  }: MediaHeaderProps) {
  ```

  Inside the right-hand `<div className="flex-1 min-w-0">`, immediately after the `<h2>` title row, add a flex container so the delete button can sit on the far right of the header. Easier: put the button at the top-right of the title row by wrapping the title in a flex layout:

  ```tsx
  <div className="flex items-start gap-2">
    <h2 className="text-lg font-bold text-ctp-text truncate flex-1">
      {detail.canonicalTitle}
    </h2>
    {onDeleteEpisode && (
      <button
        type="button"
        onClick={onDeleteEpisode}
        className="text-xs text-ctp-red/80 hover:text-ctp-red px-2 py-1 rounded hover:bg-ctp-red/10 transition-colors shrink-0"
        title="Delete this episode and all its sessions"
      >
        Delete Episode
      </button>
    )}
  </div>
  ```

- [ ] **Step 4: Wire `onDeleteEpisode` in `MediaDetailView.tsx`**

  Add a handler near the existing `handleDeleteSession`:

  ```tsx
  const handleDeleteEpisode = async () => {
    const title = data.detail.canonicalTitle;
    if (!confirmEpisodeDelete(title)) return;
    setDeleteError(null);
    try {
      await apiClient.deleteVideo(videoId);
      onBack();
    } catch (err) {
      setDeleteError(err instanceof Error ? err.message : 'Failed to delete episode.');
    }
  };
  ```

  Add `confirmEpisodeDelete` to the existing `delete-confirm` import line.

  Pass the handler down: `<MediaHeader detail={detail} ... onDeleteEpisode={handleDeleteEpisode} />`.

- [ ] **Step 5: Run the test to verify it passes**

  Run: `bun test stats/src/components/library/MediaDetailView.test.tsx -t 'delete'`
  Expected: PASS.

- [ ] **Step 6: Add `refresh` to `useMediaLibrary`**

  In `stats/src/hooks/useMediaLibrary.ts`, hoist the `load` function out of `useEffect` using `useCallback`, and return a `refresh` function:

  ```tsx
  import { useState, useEffect, useCallback } from 'react';

  export function useMediaLibrary() {
    const [media, setMedia] = useState<MediaLibraryItem[]>([]);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState<string | null>(null);
    const [version, setVersion] = useState(0);

    const refresh = useCallback(() => setVersion((v) => v + 1), []);

    useEffect(() => {
      let cancelled = false;
      let retryCount = 0;
      let retryTimer: ReturnType<typeof setTimeout> | null = null;

      const load = (isInitial = false) => {
        if (isInitial) {
          setLoading(true);
          setError(null);
        }
        getStatsClient()
          .getMediaLibrary()
          .then((rows) => {
            if (cancelled) return;
            setMedia(rows);
            if (shouldRefreshMediaLibraryRows(rows) && retryCount < MEDIA_LIBRARY_MAX_RETRIES) {
              retryCount += 1;
              retryTimer = setTimeout(() => {
                retryTimer = null;
                load(false);
              }, MEDIA_LIBRARY_REFRESH_DELAY_MS);
            }
          })
          .catch((err: Error) => {
            if (cancelled) return;
            setError(err.message);
          })
          .finally(() => {
            if (cancelled || !isInitial) return;
            setLoading(false);
          });
      };

      load(true);
      return () => {
        cancelled = true;
        if (retryTimer) {
          clearTimeout(retryTimer);
        }
      };
    }, [version]);

    return { media, loading, error, refresh };
  }
  ```

- [ ] **Step 7: Add a focused test for `refresh`**

  In `stats/src/hooks/useMediaLibrary.test.ts`, add a test that:
  - Mounts the hook with `renderHook` from `@testing-library/react`.
  - Asserts `getMediaLibrary` was called once.
  - Calls `result.current.refresh()` inside `act`.
  - Asserts `getMediaLibrary` was called twice.

  If the file doesn't have `renderHook` patterns, mirror whichever helper the existing tests use. Look at the existing test file first.

- [ ] **Step 8: Wire `refresh` from `LibraryTab.tsx`**

  In `stats/src/components/library/LibraryTab.tsx`:

  ```tsx
  const { media, loading, error, refresh } = useMediaLibrary();
  ```

  And update the early-return that mounts the detail view:

  ```tsx
  if (selectedVideoId !== null) {
    return (
      <MediaDetailView
        videoId={selectedVideoId}
        onBack={() => {
          setSelectedVideoId(null);
          refresh();
        }}
      />
    );
  }
  ```

- [ ] **Step 9: Run the new tests**

  Run: `bun test stats/src/hooks/useMediaLibrary.test.ts && bun test stats/src/components/library/MediaDetailView.test.tsx`
  Expected: PASS.

- [ ] **Step 10: Typecheck**

  Run: `bun run typecheck:stats`
  Expected: succeeds.

- [ ] **Step 11: Commit**

  ```bash
  git add stats/src/components/library/MediaHeader.tsx \
          stats/src/components/library/MediaDetailView.tsx \
          stats/src/components/library/MediaDetailView.test.tsx \
          stats/src/components/library/LibraryTab.tsx \
          stats/src/hooks/useMediaLibrary.ts \
          stats/src/hooks/useMediaLibrary.test.ts
  git commit -m "feat(stats): delete episode from library detail view"
  ```

---

## Task 7: Library — collapsible series groups

**Files:**
- Modify: `stats/src/components/library/LibraryTab.tsx`
- Test: create `stats/src/components/library/LibraryTab.test.tsx`

- [ ] **Step 1: Write the failing tests**

  Create `stats/src/components/library/LibraryTab.test.tsx`. Mirror the import/mocking pattern from `MediaDetailView.test.tsx`. Stub `useMediaLibrary` to return:
  - One group with three videos (multi-video series).
  - One group with one video (singleton).

  Add three tests:

  ```tsx
  it('renders the multi-video group collapsed by default', async () => {
    // Render LibraryTab with stubbed hook.
    // Assert: the group header is visible.
    // Assert: the three video MediaCards are NOT in the DOM (collapsed).
  });

  it('renders the single-video group expanded by default', async () => {
    // Assert: the singleton's MediaCard IS in the DOM.
  });

  it('toggles the collapsed group when its header is clicked', async () => {
    // Click the multi-video group header.
    // Assert: the three MediaCards now appear.
    // Click again.
    // Assert: they disappear.
  });
  ```

  How to identify cards: each `MediaCard` should expose its title via the cover image alt text or a title element. Use `screen.queryAllByText(<title>)` to count them.

- [ ] **Step 2: Run the failing tests**

  Run: `bun test stats/src/components/library/LibraryTab.test.tsx`
  Expected: FAIL — current `LibraryTab` always shows all cards.

- [ ] **Step 3: Add collapsible state and toggle to `LibraryTab.tsx`**

  Modify imports:
  ```tsx
  import { useState, useMemo, useCallback } from 'react';
  ```

  Inside the component, after the existing `useState` calls:
  ```tsx
  const [collapsedGroups, setCollapsedGroups] = useState<Set<string>>(() => new Set());

  // When grouped data changes, default-collapse groups with >1 video.
  // We do this declaratively in a useMemo to keep state derived.
  const effectiveCollapsed = useMemo(() => {
    const next = new Set(collapsedGroups);
    for (const group of grouped) {
      // Only auto-collapse on first encounter; if user has interacted, leave alone.
      // We do this by tracking which keys we've seen via a ref. Simpler approach:
      // initialize on mount via useEffect below.
    }
    return next;
  }, [collapsedGroups, grouped]);
  ```

  Actually, the cleanest pattern is **initialize once on first data load via `useEffect`**:
  ```tsx
  const [collapsedGroups, setCollapsedGroups] = useState<Set<string>>(() => new Set());
  const [hasInitializedCollapsed, setHasInitializedCollapsed] = useState(false);

  useEffect(() => {
    if (hasInitializedCollapsed || grouped.length === 0) return;
    const initial = new Set<string>();
    for (const group of grouped) {
      if (group.items.length > 1) initial.add(group.key);
    }
    setCollapsedGroups(initial);
    setHasInitializedCollapsed(true);
  }, [grouped, hasInitializedCollapsed]);

  const toggleGroup = useCallback((key: string) => {
    setCollapsedGroups((prev) => {
      const next = new Set(prev);
      if (next.has(key)) {
        next.delete(key);
      } else {
        next.add(key);
      }
      return next;
    });
  }, []);
  ```

  Don't forget to add `useEffect` to the import line.

- [ ] **Step 4: Update the group rendering**

  Replace the section block (around line 64-115) so the header is a `<button>`:

  ```tsx
  {grouped.map((group) => {
    const isCollapsed = collapsedGroups.has(group.key);
    const isSingleVideo = group.items.length === 1;
    return (
      <section
        key={group.key}
        className="rounded-2xl border border-ctp-surface1 bg-ctp-surface0/70 overflow-hidden"
      >
        <button
          type="button"
          onClick={() => !isSingleVideo && toggleGroup(group.key)}
          aria-expanded={!isCollapsed}
          aria-controls={`group-body-${group.key}`}
          disabled={isSingleVideo}
          className={`w-full flex items-center gap-4 p-4 border-b border-ctp-surface1 bg-ctp-base/40 text-left ${
            isSingleVideo ? '' : 'hover:bg-ctp-base/60 transition-colors cursor-pointer'
          }`}
        >
          {!isSingleVideo && (
            <span
              aria-hidden="true"
              className={`text-xs text-ctp-overlay2 transition-transform shrink-0 ${
                isCollapsed ? '' : 'rotate-90'
              }`}
            >
              {'\u25B6'}
            </span>
          )}
          <CoverImage
            videoId={group.items[0]!.videoId}
            title={group.title}
            src={group.imageUrl}
            className="w-16 h-16 rounded-2xl shrink-0"
          />
          <div className="min-w-0 flex-1">
            <div className="flex items-center gap-2">
              <h3 className="text-base font-semibold text-ctp-text truncate">
                {group.title}
              </h3>
            </div>
            {group.subtitle ? (
              <div className="text-xs text-ctp-overlay1 truncate mt-1">{group.subtitle}</div>
            ) : null}
            <div className="text-xs text-ctp-overlay2 mt-2">
              {group.items.length} video{group.items.length !== 1 ? 's' : ''} ·{' '}
              {formatDuration(group.totalActiveMs)} · {formatNumber(group.totalCards)} cards
            </div>
          </div>
        </button>
        {!isCollapsed && (
          <div id={`group-body-${group.key}`} className="p-4">
            <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 gap-4">
              {group.items.map((item) => (
                <MediaCard
                  key={item.videoId}
                  item={item}
                  onClick={() => setSelectedVideoId(item.videoId)}
                />
              ))}
            </div>
          </div>
        )}
      </section>
    );
  })}
  ```

  **Watch out:** the previous header had a clickable `<a>` for the channel URL. Wrapping the whole header in a `<button>` makes nested anchors invalid. The simplest fix: drop the channel URL link from inside the header (it's still reachable from the individual `MediaCard`s), or move it to a separate row outside the button. Choose the first — minimum visual disruption.

- [ ] **Step 5: Run the tests to verify they pass**

  Run: `bun test stats/src/components/library/LibraryTab.test.tsx`
  Expected: PASS.

- [ ] **Step 6: Typecheck**

  Run: `bun run typecheck:stats`
  Expected: succeeds.

- [ ] **Step 7: Commit**

  ```bash
  git add stats/src/components/library/LibraryTab.tsx \
          stats/src/components/library/LibraryTab.test.tsx
  git commit -m "feat(stats): collapsible series groups in library tab"
  ```

---

## Task 8: Session grouping helper

**Files:**
- Create: `stats/src/lib/session-grouping.ts`
- Create: `stats/src/lib/session-grouping.test.ts`

- [ ] **Step 1: Write the failing tests**

  Create `stats/src/lib/session-grouping.test.ts`:

  ```ts
  import { describe, it, expect } from 'bun:test';
  import { groupSessionsByVideo } from './session-grouping';
  import type { SessionSummary } from '../types/stats';

  function makeSession(over: Partial<SessionSummary>): SessionSummary {
    return {
      sessionId: 1,
      videoId: 100,
      canonicalTitle: 'Episode 1',
      animeTitle: 'Show',
      startedAtMs: 1_000_000,
      activeWatchedMs: 60_000,
      cardsMined: 1,
      linesSeen: 10,
      lookupCount: 5,
      lookupHits: 3,
      knownWordsSeen: 5,
      // Add any other required fields by reading types/stats.ts.
      ...over,
    } as SessionSummary;
  }

  describe('groupSessionsByVideo', () => {
    it('returns an empty array for empty input', () => {
      expect(groupSessionsByVideo([])).toEqual([]);
    });

    it('emits a singleton bucket for unique videoIds', () => {
      const a = makeSession({ sessionId: 1, videoId: 100 });
      const b = makeSession({ sessionId: 2, videoId: 200 });
      const buckets = groupSessionsByVideo([a, b]);
      expect(buckets).toHaveLength(2);
      expect(buckets[0]!.sessions).toHaveLength(1);
      expect(buckets[1]!.sessions).toHaveLength(1);
    });

    it('combines multiple sessions sharing a videoId into one bucket with summed totals', () => {
      const a = makeSession({
        sessionId: 1,
        videoId: 100,
        startedAtMs: 1_000_000,
        activeWatchedMs: 60_000,
        cardsMined: 2,
      });
      const b = makeSession({
        sessionId: 2,
        videoId: 100,
        startedAtMs: 2_000_000,
        activeWatchedMs: 120_000,
        cardsMined: 3,
      });
      const buckets = groupSessionsByVideo([a, b]);
      expect(buckets).toHaveLength(1);
      const bucket = buckets[0]!;
      expect(bucket.sessions).toHaveLength(2);
      expect(bucket.totalActiveMs).toBe(180_000);
      expect(bucket.totalCardsMined).toBe(5);
      // Representative is the most-recent session.
      expect(bucket.representativeSession.sessionId).toBe(2);
    });

    it('treats sessions with null/missing videoId as singletons keyed by sessionId', () => {
      const a = makeSession({ sessionId: 1, videoId: null as unknown as number });
      const b = makeSession({ sessionId: 2, videoId: null as unknown as number });
      const buckets = groupSessionsByVideo([a, b]);
      expect(buckets).toHaveLength(2);
      expect(buckets[0]!.key).toContain('1');
      expect(buckets[1]!.key).toContain('2');
    });
  });
  ```

- [ ] **Step 2: Run the failing tests**

  Run: `bun test stats/src/lib/session-grouping.test.ts`
  Expected: FAIL — module does not exist.

- [ ] **Step 3: Implement the helper**

  Create `stats/src/lib/session-grouping.ts`:

  ```ts
  import type { SessionSummary } from '../types/stats';

  export interface SessionBucket {
    key: string;
    videoId: number | null;
    sessions: SessionSummary[];
    totalActiveMs: number;
    totalCardsMined: number;
    representativeSession: SessionSummary;
  }

  export function groupSessionsByVideo(sessions: SessionSummary[]): SessionBucket[] {
    const byVideo = new Map<string, SessionSummary[]>();

    for (const session of sessions) {
      const hasVideoId =
        typeof session.videoId === 'number' && Number.isFinite(session.videoId) && session.videoId > 0;
      const key = hasVideoId ? `v-${session.videoId}` : `s-${session.sessionId}`;
      const existing = byVideo.get(key);
      if (existing) {
        existing.push(session);
      } else {
        byVideo.set(key, [session]);
      }
    }

    const buckets: SessionBucket[] = [];
    for (const [key, group] of byVideo) {
      const sorted = [...group].sort((a, b) => b.startedAtMs - a.startedAtMs);
      const representative = sorted[0]!;
      buckets.push({
        key,
        videoId:
          typeof representative.videoId === 'number' && representative.videoId > 0
            ? representative.videoId
            : null,
        sessions: sorted,
        totalActiveMs: sorted.reduce((sum, s) => sum + s.activeWatchedMs, 0),
        totalCardsMined: sorted.reduce((sum, s) => sum + s.cardsMined, 0),
        representativeSession: representative,
      });
    }

    // Preserve insertion order — `byVideo` already keeps it (Map insertion order).
    return buckets;
  }
  ```

- [ ] **Step 4: Run the tests to verify they pass**

  Run: `bun test stats/src/lib/session-grouping.test.ts`
  Expected: PASS.

- [ ] **Step 5: Typecheck**

  Run: `bun run typecheck:stats`
  Expected: succeeds.

- [ ] **Step 6: Commit**

  ```bash
  git add stats/src/lib/session-grouping.ts stats/src/lib/session-grouping.test.ts
  git commit -m "feat(stats): add groupSessionsByVideo helper for episode rollups"
  ```

---

## Task 9: Sessions tab — episode rollup UI

**Files:**
- Modify: `stats/src/components/sessions/SessionsTab.tsx`
- Modify: `stats/src/lib/delete-confirm.ts` (add `confirmBucketDelete`)
- Modify: `stats/src/lib/delete-confirm.test.ts`
- Test: extend `stats/src/components/sessions/SessionsTab.test.tsx` if it exists; otherwise add a focused integration test on the new rollup behavior.

- [ ] **Step 1: Add `confirmBucketDelete` with a failing test**

  In `stats/src/lib/delete-confirm.test.ts`, add:

  ```ts
  it('confirmBucketDelete asks about merging multiple sessions of the same episode', () => {
    // mock globalThis.confirm to capture the prompt and return true
    const calls: string[] = [];
    const original = globalThis.confirm;
    globalThis.confirm = ((msg: string) => {
      calls.push(msg);
      return true;
    }) as typeof globalThis.confirm;
    try {
      expect(confirmBucketDelete('My Episode', 3)).toBe(true);
      expect(calls[0]).toContain('3');
      expect(calls[0]).toContain('My Episode');
    } finally {
      globalThis.confirm = original;
    }
  });
  ```

  Update the import line to also import `confirmBucketDelete`.

- [ ] **Step 2: Run the failing test**

  Run: `bun test stats/src/lib/delete-confirm.test.ts -t 'confirmBucketDelete'`
  Expected: FAIL — function doesn't exist.

- [ ] **Step 3: Add the helper**

  Append to `stats/src/lib/delete-confirm.ts`:

  ```ts
  export function confirmBucketDelete(title: string, count: number): boolean {
    return globalThis.confirm(
      `Delete all ${count} session${count === 1 ? '' : 's'} of "${title}" from this day?`,
    );
  }
  ```

- [ ] **Step 4: Re-run the test**

  Run: `bun test stats/src/lib/delete-confirm.test.ts`
  Expected: PASS.

- [ ] **Step 5: Add a failing test for the bucket UI**

  In a new or extended `stats/src/components/sessions/SessionsTab.test.tsx`, add a test that:
  - Stubs `useSessions` to return three sessions on the same day, two of which share a `videoId`.
  - Renders `<SessionsTab />`.
  - Asserts the page contains a bucket header for the shared-video pair (e.g. text matching `2 sessions`).
  - Asserts the singleton session's title appears once.
  - Clicks the bucket header and verifies the underlying two sessions become visible.

- [ ] **Step 6: Run the failing test**

  Run: `bun test stats/src/components/sessions/SessionsTab.test.tsx -t 'rollup'`
  Expected: FAIL — current behavior renders three flat rows.

- [ ] **Step 7: Restructure `SessionsTab.tsx` to use buckets**

  At the top of the component, import the helper:

  ```tsx
  import { groupSessionsByVideo, type SessionBucket } from '../../lib/session-grouping';
  import { confirmBucketDelete } from '../../lib/delete-confirm';
  ```

  Add a second expanded-state Set keyed by bucket key:

  ```tsx
  const [expandedBuckets, setExpandedBuckets] = useState<Set<string>>(new Set());

  const toggleBucket = (key: string) => {
    setExpandedBuckets((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };
  ```

  Replace the inner day-group loop. Instead of mapping `daySessions.map(...)` directly, run them through `groupSessionsByVideo` and render each bucket. Buckets with one session keep the existing `SessionRow` rendering. Buckets with multiple sessions render a `<SessionBucketRow>` (a small inline component or a JSX block — keep it inline if the file isn't getting too long).

  Skeleton:

  ```tsx
  {Array.from(groups.entries()).map(([dayLabel, daySessions]) => {
    const buckets = groupSessionsByVideo(daySessions);
    return (
      <div key={dayLabel}>
        <div className="flex items-center gap-3 mb-2">
          <h3 className="text-xs font-semibold text-ctp-overlay2 uppercase tracking-widest shrink-0">
            {dayLabel}
          </h3>
          <div className="flex-1 h-px bg-gradient-to-r from-ctp-surface1 to-transparent" />
        </div>
        <div className="space-y-2">
          {buckets.map((bucket) => {
            if (bucket.sessions.length === 1) {
              const s = bucket.sessions[0]!;
              const detailsId = `session-details-${s.sessionId}`;
              return (
                <div key={bucket.key}>
                  <SessionRow
                    session={s}
                    isExpanded={expandedId === s.sessionId}
                    detailsId={detailsId}
                    onToggle={() => setExpandedId(expandedId === s.sessionId ? null : s.sessionId)}
                    onDelete={() => void handleDeleteSession(s)}
                    deleteDisabled={deletingSessionId === s.sessionId}
                    onNavigateToMediaDetail={onNavigateToMediaDetail}
                  />
                  {expandedId === s.sessionId && (
                    <div id={detailsId}>
                      <SessionDetail session={s} />
                    </div>
                  )}
                </div>
              );
            }
            const isOpen = expandedBuckets.has(bucket.key);
            return (
              <div key={bucket.key} className="rounded-lg border border-ctp-surface1 bg-ctp-surface0/40">
                <button
                  type="button"
                  onClick={() => toggleBucket(bucket.key)}
                  aria-expanded={isOpen}
                  className="w-full flex items-center gap-3 px-3 py-2 text-left hover:bg-ctp-surface0/70 transition-colors"
                >
                  <span
                    aria-hidden="true"
                    className={`text-xs text-ctp-overlay2 transition-transform ${isOpen ? 'rotate-90' : ''}`}
                  >
                    {'\u25B6'}
                  </span>
                  <div className="min-w-0 flex-1">
                    <div className="text-sm text-ctp-text truncate">
                      {bucket.representativeSession.canonicalTitle ?? 'Unknown Episode'}
                    </div>
                    <div className="text-xs text-ctp-overlay2">
                      {bucket.sessions.length} sessions ·{' '}
                      {formatDuration(bucket.totalActiveMs)} ·{' '}
                      {bucket.totalCardsMined} cards
                    </div>
                  </div>
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      void handleDeleteBucket(bucket);
                    }}
                    className="text-[10px] text-ctp-red/70 hover:text-ctp-red px-1.5 py-0.5 rounded hover:bg-ctp-red/10 transition-colors"
                    title="Delete all sessions in this group"
                  >
                    Delete
                  </button>
                </button>
                {isOpen && (
                  <div className="pl-8 pr-2 pb-2 space-y-2">
                    {bucket.sessions.map((s) => {
                      const detailsId = `session-details-${s.sessionId}`;
                      return (
                        <div key={s.sessionId}>
                          <SessionRow
                            session={s}
                            isExpanded={expandedId === s.sessionId}
                            detailsId={detailsId}
                            onToggle={() =>
                              setExpandedId(expandedId === s.sessionId ? null : s.sessionId)
                            }
                            onDelete={() => void handleDeleteSession(s)}
                            deleteDisabled={deletingSessionId === s.sessionId}
                            onNavigateToMediaDetail={onNavigateToMediaDetail}
                          />
                          {expandedId === s.sessionId && (
                            <div id={detailsId}>
                              <SessionDetail session={s} />
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>
    );
  })}
  ```

  **Note on nested buttons:** the bucket header is a `<button>` and contains a "Delete" `<button>`. HTML disallows nested buttons. Switch the outer element to a `<div role="button" tabIndex={0} onClick={...} onKeyDown={...}>` instead, OR put the delete button in a wrapping flex container *outside* the toggle button. Pick the second option — it's accessible without role gymnastics:

  ```tsx
  <div className="flex items-center">
    <button type="button" onClick={() => toggleBucket(bucket.key)} ...>
      ...
    </button>
    <button type="button" onClick={() => void handleDeleteBucket(bucket)} ...>
      Delete
    </button>
  </div>
  ```

  Use that pattern in the actual implementation. The skeleton above shows the *intent*; the final code must have sibling buttons, not nested ones.

  Add `handleDeleteBucket`:

  ```tsx
  const handleDeleteBucket = async (bucket: SessionBucket) => {
    const title = bucket.representativeSession.canonicalTitle ?? 'this episode';
    if (!confirmBucketDelete(title, bucket.sessions.length)) return;
    setDeleteError(null);
    const ids = bucket.sessions.map((s) => s.sessionId);
    try {
      await apiClient.deleteSessions(ids);
      const idSet = new Set(ids);
      setVisibleSessions((prev) => prev.filter((s) => !idSet.has(s.sessionId)));
    } catch (err) {
      setDeleteError(err instanceof Error ? err.message : 'Failed to delete sessions.');
    }
  };
  ```

  Add the `formatDuration` import at the top of the file if not present.

- [ ] **Step 8: Run the bucket test**

  Run: `bun test stats/src/components/sessions/SessionsTab.test.tsx -t 'rollup'`
  Expected: PASS.

- [ ] **Step 9: Run all sessions tests**

  Run: `bun test stats/src/components/sessions/`
  Expected: PASS.

- [ ] **Step 10: Apply the same rollup to `MediaSessionList.tsx`**

  Read `stats/src/components/library/MediaSessionList.tsx` first. Inside a single video's detail view, all sessions share the same `videoId`, so `groupSessionsByVideo` would always produce one giant bucket. **That's wrong for this view.** Skip the bucket rendering here entirely — `MediaSessionList` still groups by day only. Document this in the commit message: "Rollup intentionally not applied in MediaSessionList because the view is already filtered to a single video."

- [ ] **Step 11: Typecheck**

  Run: `bun run typecheck:stats`
  Expected: succeeds.

- [ ] **Step 12: Commit**

  ```bash
  git add stats/src/components/sessions/SessionsTab.tsx \
          stats/src/components/sessions/SessionsTab.test.tsx \
          stats/src/lib/delete-confirm.ts stats/src/lib/delete-confirm.test.ts
  git commit -m "feat(stats): roll up same-episode sessions within a day"
  ```

---

## Task 10: Chart clarity pass

**Files:**
- Modify: `stats/src/lib/chart-theme.ts`
- Modify: `stats/src/components/trends/TrendChart.tsx`
- Modify: `stats/src/components/trends/StackedTrendChart.tsx`
- Modify: `stats/src/components/overview/WatchTimeChart.tsx`
- Test: create `stats/src/lib/chart-theme.test.ts` if not present

- [ ] **Step 1: Extend `chart-theme.ts`**

  Replace the file contents with:

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

- [ ] **Step 2: Add a snapshot/value test for the new constants**

  Create `stats/src/lib/chart-theme.test.ts`:

  ```ts
  import { describe, it, expect } from 'bun:test';
  import { CHART_THEME, CHART_DEFAULTS, TOOLTIP_CONTENT_STYLE } from './chart-theme';

  describe('chart-theme', () => {
    it('exposes a grid color', () => {
      expect(CHART_THEME.grid).toBe('#494d64');
    });

    it('uses 11px ticks for legibility', () => {
      expect(CHART_DEFAULTS.tickFontSize).toBe(11);
    });

    it('builds a tooltip content style with border + background', () => {
      expect(TOOLTIP_CONTENT_STYLE.background).toBe(CHART_THEME.tooltipBg);
      expect(TOOLTIP_CONTENT_STYLE.border).toContain(CHART_THEME.tooltipBorder);
    });
  });
  ```

- [ ] **Step 3: Run the test**

  Run: `bun test stats/src/lib/chart-theme.test.ts`
  Expected: PASS.

- [ ] **Step 4: Update `TrendChart.tsx` to use the shared theme + add gridlines**

  Replace `stats/src/components/trends/TrendChart.tsx` with:

  ```tsx
  import {
    BarChart,
    Bar,
    LineChart,
    Line,
    XAxis,
    YAxis,
    Tooltip,
    CartesianGrid,
    ResponsiveContainer,
  } from 'recharts';
  import { CHART_THEME, CHART_DEFAULTS, TOOLTIP_CONTENT_STYLE } from '../../lib/chart-theme';

  interface TrendChartProps {
    title: string;
    data: Array<{ label: string; value: number }>;
    color: string;
    type: 'bar' | 'line';
    formatter?: (value: number) => string;
    onBarClick?: (label: string) => void;
  }

  export function TrendChart({ title, data, color, type, formatter, onBarClick }: TrendChartProps) {
    const formatValue = (v: number) => (formatter ? [formatter(v), title] : [String(v), title]);

    return (
      <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
        <h3 className="text-xs font-semibold text-ctp-text mb-2">{title}</h3>
        <ResponsiveContainer width="100%" height={CHART_DEFAULTS.height}>
          {type === 'bar' ? (
            <BarChart data={data} margin={CHART_DEFAULTS.margin}>
              <CartesianGrid stroke={CHART_THEME.grid} {...CHART_DEFAULTS.grid} />
              <XAxis
                dataKey="label"
                tick={{ fontSize: CHART_DEFAULTS.tickFontSize, fill: CHART_THEME.tick }}
                axisLine={{ stroke: CHART_THEME.axisLine }}
                tickLine={false}
              />
              <YAxis
                tick={{ fontSize: CHART_DEFAULTS.tickFontSize, fill: CHART_THEME.tick }}
                axisLine={{ stroke: CHART_THEME.axisLine }}
                tickLine={false}
                width={32}
                tickFormatter={formatter}
              />
              <Tooltip contentStyle={TOOLTIP_CONTENT_STYLE} formatter={formatValue} />
              <Bar
                dataKey="value"
                fill={color}
                radius={[2, 2, 0, 0]}
                cursor={onBarClick ? 'pointer' : undefined}
                onClick={
                  onBarClick ? (entry: { label: string }) => onBarClick(entry.label) : undefined
                }
              />
            </BarChart>
          ) : (
            <LineChart data={data} margin={CHART_DEFAULTS.margin}>
              <CartesianGrid stroke={CHART_THEME.grid} {...CHART_DEFAULTS.grid} />
              <XAxis
                dataKey="label"
                tick={{ fontSize: CHART_DEFAULTS.tickFontSize, fill: CHART_THEME.tick }}
                axisLine={{ stroke: CHART_THEME.axisLine }}
                tickLine={false}
              />
              <YAxis
                tick={{ fontSize: CHART_DEFAULTS.tickFontSize, fill: CHART_THEME.tick }}
                axisLine={{ stroke: CHART_THEME.axisLine }}
                tickLine={false}
                width={32}
                tickFormatter={formatter}
              />
              <Tooltip contentStyle={TOOLTIP_CONTENT_STYLE} formatter={formatValue} />
              <Line dataKey="value" stroke={color} strokeWidth={2} dot={false} />
            </LineChart>
          )}
        </ResponsiveContainer>
      </div>
    );
  }
  ```

- [ ] **Step 5: Update `StackedTrendChart.tsx` and `WatchTimeChart.tsx`**

  Open each file. For each chart container, apply the same recipe:
  1. Import `CartesianGrid` from `recharts`.
  2. Import `CHART_THEME`, `CHART_DEFAULTS`, `TOOLTIP_CONTENT_STYLE` from `'../../lib/chart-theme'`.
  3. Insert `<CartesianGrid stroke={CHART_THEME.grid} {...CHART_DEFAULTS.grid} />` as the first child of `<BarChart>`/`<LineChart>`.
  4. Bump `<XAxis tick fontSize>` and `<YAxis tick fontSize>` to `CHART_DEFAULTS.tickFontSize`.
  5. Add `axisLine={{ stroke: CHART_THEME.axisLine }}` to the Y axis.
  6. Replace inline tooltip styles with `contentStyle={TOOLTIP_CONTENT_STYLE}`.
  7. Bump `<ResponsiveContainer height>` from its current value to `CHART_DEFAULTS.height` (160) only if it's currently smaller than 160. Don't shrink anything.

  If either file already exposes formatter props for the Y axis, also pass `tickFormatter={formatter}` to `YAxis` so the unit suffix shows up.

- [ ] **Step 6: Re-run the chart-theme test plus typecheck**

  Run: `bun test stats/src/lib/chart-theme.test.ts && bun run typecheck:stats`
  Expected: PASS + clean.

- [ ] **Step 7: Sanity-check the overview tab still mounts**

  Run: `bun run build:stats`
  Expected: succeeds.

- [ ] **Step 8: Commit**

  ```bash
  git add stats/src/lib/chart-theme.ts stats/src/lib/chart-theme.test.ts \
          stats/src/components/trends/TrendChart.tsx \
          stats/src/components/trends/StackedTrendChart.tsx \
          stats/src/components/overview/WatchTimeChart.tsx
  git commit -m "feat(stats): unify chart theme and add gridlines for legibility"
  ```

---

## Task 11: Changelog fragment

**Files:**
- Create: `changes/2026-04-09-stats-dashboard-feedback-pass.md`

- [ ] **Step 1: Read the existing changelog format**

  Run: `ls changes/ | head -5 && cat changes/$(ls changes/ | head -1)`
  Mirror that format exactly.

- [ ] **Step 2: Write the fragment**

  Create `changes/2026-04-09-stats-dashboard-feedback-pass.md` with content like:

  ```markdown
  ---
  type: feature
  scope: stats
  ---

  Stats dashboard polish:

  - Library now collapses multi-episode series under a clickable header.
  - Sessions tab rolls up multiple sessions of the same episode within a day.
  - Trends gain a 365d range option.
  - Episodes can be deleted directly from the library detail view.
  - Top 50 vocabulary tightens word/reading spacing.
  - Cards deleted from Anki no longer appear in the episode detail card list.
  - Trend and watch-time charts gain horizontal gridlines, larger ticks, and a shared theme.
  ```

  Adjust frontmatter keys/values to match whatever existing fragments use.

- [ ] **Step 3: Validate**

  Run: `bun run changelog:lint && bun run changelog:pr-check`
  Expected: PASS.

- [ ] **Step 4: Commit**

  ```bash
  git add changes/2026-04-09-stats-dashboard-feedback-pass.md
  git commit -m "docs: add changelog fragment for stats dashboard feedback pass"
  ```

---

## Final verification gate

Run the project's standard handoff gate:

- [ ] `bun run typecheck`
- [ ] `bun run typecheck:stats`
- [ ] `bun run test:fast`
- [ ] `bun run test:env`
- [ ] `bun run test:runtime:compat`
- [ ] `bun run build`
- [ ] `bun run test:smoke:dist`
- [ ] `bun run format:check:src`
- [ ] `bun run changelog:lint`
- [ ] `bun run changelog:pr-check`

If any of those fail, fix the underlying issue and create a new commit (do NOT amend earlier task commits — keep the per-task history clean).

Then push the branch and open the PR. Suggested PR title:

```
Stats dashboard polish: collapsible library, session rollups, 365d trends, chart legibility, episode delete
```

Body should link to the spec at `docs/superpowers/specs/2026-04-09-stats-dashboard-feedback-pass-design.md` and summarize each task.

---

## Risk callouts (for the implementing agent)

- **Anki note-info loading-state guard (Task 5):** double-check the test case for the brief window before `ankiNotesInfo` resolves. Hiding everything during that window would be a regression.
- **Nested button trap (Task 9):** the bucket header must place the toggle button and the delete button as siblings, not nested. Final code must use sibling buttons; the skeleton in the plan flags this.
- **MediaSessionList (Task 9):** rollup is intentionally not applied there. Don't forget the commit message note.
- **`useMediaLibrary` retry behavior (Task 6):** the existing hook auto-refetches when youtube metadata is missing. The new `refresh()` must not break that loop. The `[version]` dependency on the existing `useEffect` triggers a brand-new mount of the inner closure each call, which resets `retryCount` — that's the intended behavior.
- **`bun test` resolves test files relative to repo root.** Always run from `/Users/sudacode/projects/japanese/SubMiner` (the worktree root), not from `stats/`.
- **No file in this plan grows past ~250 lines after edits.** If a file does, that's a signal to extract — flag it on the way through.
