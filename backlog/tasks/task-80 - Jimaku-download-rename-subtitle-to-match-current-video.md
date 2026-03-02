---
id: TASK-80
title: 'Jimaku download: rename subtitle to current video basename'
status: Done
assignee: []
created_date: '2026-03-01 14:17'
updated_date: '2026-03-01 14:19'
labels: []
dependencies: []
priority: medium
ordinal: 11000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

When user selects a Jimaku subtitle, save subtitle with filename derived from currently playing media filename instead of Jimaku release filename.

Example:
- Current media: `anime.mkv`
- Downloaded subtitle extension: `.srt`
- Saved subtitle path: `anime.ja.srt`

Scope:
- Apply in Jimaku download IPC path before writing file.
- Preserve collision-avoidance behavior (suffix with jimaku entry id/counter when target exists).
- Keep mpv load flow unchanged except using renamed path.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Jimaku subtitle destination name uses current media basename plus `.ja` and subtitle extension.
- [x] #2 Existing duplicate filename conflict handling still works.
- [x] #3 Regression tests cover renamed destination path behavior.

<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Jimaku download path generation now derives subtitle filename from currently playing media basename and keeps subtitle extension from Jimaku file (`anime.mkv` + `.srt` => `anime.ja.srt`). Added pure helper `buildJimakuSubtitleFilenameFromMediaPath` and routed IPC download flow through it before existing duplicate-path conflict handling. Added regression tests for local path, missing extension fallback, and remote URL media paths.

<!-- SECTION:FINAL_SUMMARY:END -->
