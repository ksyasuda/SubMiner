---
id: TASK-108
title: Add AniSkip intro skip markers and OSD skip button in mpv plugin
status: In Progress
assignee:
  - codex
created_date: '2026-02-22 08:05'
updated_date: '2026-02-22 09:26'
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
- [ ] #1 Plugin calls AniSkip API and handles missing data gracefully.
- [ ] #2 Intro marker/chapter is visible in mpv when OP range exists.
- [ ] #3 OSD skip button appears only while inside OP range.
- [ ] #4 Clicking/activating button seeks to OP end.
- [ ] #5 Docs/config include new options + script message.
<!-- AC:END -->

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
<!-- SECTION:NOTES:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Focused plugin smoke validation in mpv.
- [ ] #2 Lua parse/load check passes in local environment.
- [ ] #3 Task notes capture fallback behavior.
<!-- DOD:END -->
