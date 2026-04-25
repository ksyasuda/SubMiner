---
id: TASK-291
title: Add manual AniList selection for character dictionary resolution
status: Done
assignee:
  - '@codex'
created_date: '2026-04-25 21:29'
updated_date: '2026-04-25 22:51'
labels:
  - dictionary
  - anilist
  - cli
  - ui
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add CLI and in-app UI support for correcting character dictionary anime resolution when guessit or AniList search picks the wrong series. Manual selections must apply to the whole detected series, persist across episodes, and replace stale incorrect entries in the auto-sync merged character dictionary state. Do not add tray UI. Known regression case: `Re - ZERO, Starting Life in Another World (2016) - S01E01 - - The End of the Beginning and the Beginning of the End [v2 Bluray-1080p Proper][10bit][x265][FLAC 2.0][EN+JA]-SCY.mkv` previously resolved to `10607 - Rerere no Tensai Bakabon`; it should be correctable to the chosen Re:ZERO AniList media and then reused for later files in that series.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 CLI can show AniList candidate matches for the current/target media and set a manual character-dictionary AniList override.
- [x] #2 In-app UI can show the current character-dictionary match, candidate matches, and apply an override without adding tray controls.
- [x] #3 Persisted overrides are keyed at series scope so all later episodes in the same series reuse the selected AniList media.
- [x] #4 Applying an override clears stale guess state, replaces the old incorrect active media entry in auto-sync state, rebuilds/imports the merged character dictionary, and refreshes subtitle dictionary usage.
- [x] #5 Regression tests cover the Re:ZERO filename, override reuse, stale active-media replacement, CLI handling, and IPC/UI contract behavior.
- [x] #6 Docs are updated for manual character dictionary anime selection.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a focused override/resolution layer for character dictionary media selection: derive a stable series key from filename/guessit data, persist manual AniList media overrides under user data, and expose AniList candidate search helpers.
2. Update character dictionary snapshot resolution to check manual overrides before guessit-derived AniList search, and update auto-sync so applying an override removes stale incorrect media IDs and rebuilds/imports the merged dictionary.
3. Extend CLI with commands to list candidates and set an override for current or target media.
4. Extend existing in-app settings UI via IPC/preload contracts: show current match/candidates and let user apply an override. No tray controls.
5. Use TDD: add failing regressions first for Re:ZERO parsing/override behavior, auto-sync replacement, CLI handling, IPC contract, and UI state; then implement.
6. Update docs-site/manual docs for manual character dictionary anime selection, launcher usage, and the default `Ctrl+Alt+A` modal shortcut, then run focused tests and broader gates as time permits.
<!-- SECTION:PLAN:END -->
