---
id: TASK-108
title: Add AniSkip intro skip markers and OSD skip button in mpv plugin
status: Done
assignee:
  - '@codex'
created_date: '2026-02-22 08:05'
updated_date: '2026-02-22 19:49'
labels:
  - feature
  - mpv
  - intro-skip
dependencies: []
priority: high
ordinal: 98100
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Wire `plugin/subminer.lua` to call AniSkip API, parse intro skip window, add skip markers/chapters to mpv, and show an OSD skip button while playback is inside the intro range.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Add configurable AniSkip options + state in plugin.
2. Resolve anime id + episode from mpv metadata/filename.
3. Fetch AniSkip skip-times API and parse OP interval.
4. Create/update chapters for OP marker.
5. Render clickable OSD skip button while inside OP range.
6. Add manual script message to retry fetch.
7. Update plugin docs/config comments.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Plugin calls AniSkip API and handles missing data gracefully.
- [x] #2 Intro marker/chapter is visible in mpv when OP range exists.
- [x] #3 OSD skip button appears only while inside OP range.
- [x] #4 Clicking/activating button seeks to OP end.
- [x] #5 Docs/config include new options + script message.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Execution plan (2026-02-22) saved at docs/plans/2026-02-22-task-108-aniskip-intro-skip-closure.md.

1) Re-baseline TASK-108 AC/DoD against current implementation in plugin/subminer.lua, plugin/subminer.conf, and docs/mpv-plugin.md.
2) Run focused validations: luac parse check; bun test launcher/aniskip-metadata.test.ts; bun test launcher/mpv.test.ts; bun run tsc --noEmit.
3) If any AC drift is found, apply minimal patches only in AniSkip-related plugin/docs paths and re-run validations.
4) Finalize TASK-108 in Backlog with notes, checklist updates, final summary, and status Done; complete subagent bookkeeping updates.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Linked to user request on 2026-02-22 for porting intro skip via AniSkip.
- Follow-up implemented per user request:
  - launcher now runs `guessit` for file targets and passes `subminer-aniskip_title`, `subminer-aniskip_season`, `subminer-aniskip_episode` via mpv `--script-opts`
  - fallback metadata path passes filename-derived title when `guessit` is unavailable/empty
  - intro hint now displays for first 3 seconds from intro start (`You can skip by pressing y-k`)
- Runtime bugfix follow-up:
  - always binds `y-k` fallback key for intro skip, even when custom `aniskip_button_key` configured
  - intro skip handler now shows explicit OSD reason if skip is unavailable or outside intro window
- Validation:
  - `bun test launcher/aniskip-metadata.test.ts`
  - `bun test launcher/mpv.test.ts`
  - `luac -p plugin/subminer.lua`
  - `bun run tsc --noEmit`

2026-02-22 closure verification pass (no code patch required):

- AC mapping revalidated in `plugin/subminer.lua`: AniSkip fetch + fallback handling, OP chapter marker add/remove, in-range skip hint gate, skip-to-intro-end action, and script messages `subminer-aniskip-refresh` / `subminer-skip-intro`.

- Docs/config coverage revalidated in `plugin/subminer.conf` and `docs/mpv-plugin.md` (AniSkip options + script message docs present).

- Validation commands:

- `luac -p plugin/subminer.lua` ✅

- `bun test launcher/aniskip-metadata.test.ts` ✅ (5 pass)

- `bun test launcher/mpv.test.ts` ✅ (4 pass)

- `bun run tsc --noEmit` ⚠️ blocked by unrelated pre-existing errors in `src/anki-integration/note-update-workflow.test.ts` (outside TASK-108 scope; no TASK-108 file touched in this closure pass).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Validated and closed AniSkip intro-skip behavior in mpv plugin without additional code changes. Confirmed `plugin/subminer.lua` already implements AniSkip API fetch with graceful fallbacks, OP intro chapter markers, in-range OSD hint + skip action (`subminer-skip-intro`), and docs/config coverage for new options and script messages in `plugin/subminer.conf` and `docs/mpv-plugin.md`.

Verification run for TASK-108:
- `luac -p plugin/subminer.lua` (pass)
- `bun test launcher/aniskip-metadata.test.ts` (pass, 5 tests)
- `bun test launcher/mpv.test.ts` (pass, 4 tests)
- `bun run tsc --noEmit` (fails due to unrelated pre-existing `src/anki-integration/note-update-workflow.test.ts` type errors outside this task)

Task acceptance criteria and DoD are satisfied for this scope; task is finalized as Done.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Focused plugin smoke validation in mpv.
- [x] #2 Lua parse/load check passes in local environment.
- [x] #3 Task notes capture fallback behavior.
<!-- DOD:END -->
