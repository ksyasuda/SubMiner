# Immersion Anime Metadata Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add anime-level immersion metadata, link videos to anime rows, and expose anime/season/episode query surfaces so future stats can aggregate by anime instead of only by video title.

**Architecture:** Introduce a new `imm_anime` table plus additive `imm_videos` metadata columns. Wire media ingest through a guessit-first, fallback-parser flow that always creates or reuses an anime row, stores per-video episode metadata, and upgrades provisional anime rows when AniList data becomes available. Keep existing video/session behavior compatible while adding new query surfaces in parallel.

**Tech Stack:** TypeScript, Bun, libsql SQLite, existing immersion tracker storage/query/service modules, existing AniList parser helpers (`guessit`, `parseMediaInfo`)

---

### Task 1: Add Red Tests for Schema Shape

**Files:**
- Modify: `src/core/services/immersion-tracker/storage-session.test.ts`
- Inspect: `src/core/services/immersion-tracker/storage.ts`
- Inspect: `src/core/services/immersion-tracker/types.ts`

**Step 1: Write the failing schema test**

Add assertions that `ensureSchema()` creates:
- `imm_anime`
- new `imm_videos` columns for `anime_id`, parsed filename/title, season, episode, parser source/confidence, and parse metadata

Use `PRAGMA table_info(imm_videos)` and `sqlite_master` queries instead of indirect assertions.

**Step 2: Run the targeted test to verify it fails**

Run:

```bash
bun test src/core/services/immersion-tracker/storage-session.test.ts
```

Expected: FAIL because the new table/columns do not exist yet.

**Step 3: Implement minimal schema changes**

Modify `src/core/services/immersion-tracker/storage.ts` and `src/core/services/immersion-tracker/types.ts`:
- add `imm_anime`
- add new `imm_videos` columns
- add indexes/FKs needed for anime lookup
- bump schema version for the fresh-schema path
- do not add migration/backfill logic for older DB contents

**Step 4: Re-run the targeted test**

Run:

```bash
bun test src/core/services/immersion-tracker/storage-session.test.ts
```

Expected: PASS.

**Step 5: Commit**

```bash
git add src/core/services/immersion-tracker/types.ts src/core/services/immersion-tracker/storage.ts src/core/services/immersion-tracker/storage-session.test.ts
git commit -m "feat(immersion): add anime schema and video metadata fields"
```

### Task 2: Add Red Tests for Anime Storage Identity and Upgrade Rules

**Files:**
- Modify: `src/core/services/immersion-tracker/storage-session.test.ts`
- Modify: `src/core/services/immersion-tracker/storage.ts`
- Inspect: `src/core/services/immersion-tracker/query.ts`

**Step 1: Write failing storage tests**

Add DB-backed tests for:
- creating a provisional anime row from normalized parsed title
- reusing that row for another video from the same anime
- upgrading the same row when AniList id/title metadata becomes available later
- preserving per-video season/episode values while sharing one anime row

Prefer explicit row assertions over service-level mocks.

**Step 2: Run the targeted test file to verify it fails**

Run:

```bash
bun test src/core/services/immersion-tracker/storage-session.test.ts
```

Expected: FAIL because storage helpers do not exist yet.

**Step 3: Implement minimal storage helpers**

In `src/core/services/immersion-tracker/storage.ts`, add focused helpers such as:
- normalize anime identity key from parsed title
- get/create provisional anime row
- upgrade anime row with AniList data
- update/link per-video anime metadata

Keep responsibilities narrow and composable; do not bury query logic in the service class.

**Step 4: Re-run the targeted test file**

Run:

```bash
bun test src/core/services/immersion-tracker/storage-session.test.ts
```

Expected: PASS.

**Step 5: Commit**

```bash
git add src/core/services/immersion-tracker/storage.ts src/core/services/immersion-tracker/storage-session.test.ts
git commit -m "feat(immersion): store provisional anime rows and upgrade with AniList data"
```

### Task 3: Add Red Tests for Parser Metadata Extraction

**Files:**
- Modify: `src/core/services/immersion-tracker/metadata.test.ts`
- Modify: `src/core/services/immersion-tracker/metadata.ts`
- Inspect: `src/jimaku/utils.ts`
- Inspect: `src/core/services/anilist/anilist-updater.ts`

**Step 1: Write failing parser tests**

Add tests for a helper that returns parsed anime/video metadata from a media path/title:
- uses `guessit` output first when available
- falls back to built-in parser when `guessit` throws or returns incomplete data
- preserves season/episode/title/source/confidence
- records filename/basename for per-video metadata

Use representative filenames like:
- `Little Witch Academia S02E05.mkv`
- `[SubsPlease] Frieren - 03 (1080p).mkv`

**Step 2: Run the targeted parser test file to verify it fails**

Run:

```bash
bun test src/core/services/immersion-tracker/metadata.test.ts
```

Expected: FAIL because the helper does not exist yet.

**Step 3: Implement the minimal parser helper**

In `src/core/services/immersion-tracker/metadata.ts`:
- add a focused helper that wraps guessit-first parsing
- reuse existing parser conventions instead of inventing a new format
- keep ffprobe/local media metadata behavior intact

If shared types are needed, add them in `src/core/services/immersion-tracker/types.ts`.

**Step 4: Re-run the targeted parser test**

Run:

```bash
bun test src/core/services/immersion-tracker/metadata.test.ts
```

Expected: PASS.

**Step 5: Commit**

```bash
git add src/core/services/immersion-tracker/metadata.ts src/core/services/immersion-tracker/metadata.test.ts src/core/services/immersion-tracker/types.ts
git commit -m "feat(immersion): add guessit-first anime metadata parsing helper"
```

### Task 4: Add Red Tests for Media-Change Ingest Wiring

**Files:**
- Modify: `src/core/services/immersion-tracker-service.test.ts`
- Modify: `src/core/services/immersion-tracker-service.ts`
- Inspect: `src/core/services/immersion-tracker/storage.ts`
- Inspect: `src/core/services/immersion-tracker/metadata.ts`

**Step 1: Write failing service tests**

Add focused tests showing that `handleMediaChange(...)`:
- creates/links an anime row
- stores parsed season/episode/file metadata on the active video row
- reuses the same anime row across multiple video files for the same parsed anime
- keeps working when AniList lookup is missing

Prefer DB-backed assertions after service calls rather than deep mocking.

**Step 2: Run the targeted service test to verify it fails**

Run:

```bash
bun test src/core/services/immersion-tracker-service.test.ts
```

Expected: FAIL because ingest does not yet populate anime metadata.

**Step 3: Implement the minimal service wiring**

Modify `src/core/services/immersion-tracker-service.ts` to:
- call the new parser helper during media change
- create/reuse provisional anime rows
- persist per-video metadata
- trigger AniList enrichment/upgrade only as far as current dependencies already allow

Do not refactor unrelated tracker behavior while making this pass.

**Step 4: Re-run the targeted service test**

Run:

```bash
bun test src/core/services/immersion-tracker-service.test.ts
```

Expected: PASS.

**Step 5: Commit**

```bash
git add src/core/services/immersion-tracker-service.ts src/core/services/immersion-tracker-service.test.ts
git commit -m "feat(immersion): link videos to anime metadata during media ingest"
```

### Task 5: Add Red Tests for Anime Query Surfaces

**Files:**
- Modify: `src/core/services/immersion-tracker/__tests__/query.test.ts`
- Modify: `src/core/services/immersion-tracker/query.ts`
- Modify: `src/core/services/immersion-tracker/types.ts`

**Step 1: Write failing query tests**

Add tests for new query functions such as:
- anime library summary list
- anime detail summary
- per-anime episode list or season breakdown

Seed the DB with:
- one anime with multiple episode files
- repeated sessions on one episode
- another anime for contrast

Assert grouping by `anime_id`, not by `canonical_title`.

**Step 2: Run the targeted query test to verify it fails**

Run:

```bash
bun test src/core/services/immersion-tracker/__tests__/query.test.ts
```

Expected: FAIL because the anime query functions/types do not exist yet.

**Step 3: Implement minimal query functions**

Modify `src/core/services/immersion-tracker/query.ts` and related exported types to add anime-level queries in parallel with existing video-level queries.

Keep SQL explicit and aggregation stable:
- anime totals from linked sessions/videos
- episode/season data from video-level parsed fields

**Step 4: Re-run the targeted query test**

Run:

```bash
bun test src/core/services/immersion-tracker/__tests__/query.test.ts
```

Expected: PASS.

**Step 5: Commit**

```bash
git add src/core/services/immersion-tracker/query.ts src/core/services/immersion-tracker/types.ts src/core/services/immersion-tracker/__tests__/query.test.ts
git commit -m "feat(immersion): add anime-level stats queries"
```

### Task 6: Integrate Export Surfaces and Compatibility Checks

**Files:**
- Modify: `src/core/services/immersion-tracker-service.ts`
- Modify: any stats-server or API files only if needed after query integration
- Inspect: `src/core/services/__tests__/stats-server.test.ts`
- Inspect: `stats/src/lib/dashboard-data.ts`

**Step 1: Write the smallest failing integration test if API surface changes**

Only if the service/API export surface changes, add one failing test proving the new query path is exposed correctly. If no export change is needed yet, skip straight to implementation and note the skip in the task notes.

**Step 2: Run the targeted test to verify red state**

Run only the affected test file, for example:

```bash
bun test src/core/services/__tests__/stats-server.test.ts
```

Expected: FAIL if a new API contract is required; otherwise explicitly skip.

**Step 3: Implement minimal integration**

Export new query methods through the service only where needed for the next stats consumer. Avoid prematurely reshaping the public API if current UI work is out of scope.

**Step 4: Run the targeted integration test**

Run:

```bash
bun test src/core/services/__tests__/stats-server.test.ts
```

Expected: PASS, or documented skip if no API change was needed.

**Step 5: Commit**

```bash
git add src/core/services/immersion-tracker-service.ts src/core/services/__tests__/stats-server.test.ts stats/src/lib/dashboard-data.ts
git commit -m "feat(stats): expose anime-level immersion data where needed"
```

### Task 7: Run Focused Verification and Update Docs/Task

**Files:**
- Modify: `backlog/tasks/task-169 - Add-anime-level-immersion-metadata-and-link-videos.md`
- Modify: docs only if implementation changes user-visible behavior or API expectations

**Step 1: Run the focused SQLite immersion lane**

Run:

```bash
bun run test:immersion:sqlite:src
```

Expected: PASS.

**Step 2: Run any additional required verification**

Use the repo verifier/classifier to choose broader lanes if the diff touches runtime or stats-server surfaces:

```bash
bash .agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh
bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane core
```

Escalate only if the touched files require it.

**Step 3: Update task notes and final summary**

Record:
- commands run
- pass/fail
- skipped lanes
- remaining risks

Update the task plan section if actual execution deviated.

**Step 4: Commit**

```bash
git add backlog/tasks/task-169\ -\ Add-anime-level-immersion-metadata-and-link-videos.md
git commit -m "docs(backlog): record immersion anime metadata verification"
```
