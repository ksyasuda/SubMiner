---
id: TASK-79
title: 'Jimaku modal: auto-close after successful subtitle load'
status: Done
assignee: []
created_date: '2026-03-01 13:52'
updated_date: '2026-03-01 14:06'
labels: []
dependencies: []
priority: medium
ordinal: 10000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Fix Jimaku modal UX so selecting a subtitle file closes the modal automatically once subtitle download+load succeeds.

Current behavior:

- Subtitle file downloads and loads into mpv.
- Jimaku modal remains open until manual close.

Expected behavior:

- On successful `jimakuDownloadFile` result, close modal immediately.
- Keep error behavior unchanged (stay open + show error).

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Successful subtitle file selection/download in Jimaku closes modal automatically.
- [x] #2 Existing error path keeps modal open and shows error.
- [x] #3 Regression test covers success auto-close behavior.

<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Fixed renderer Jimaku success flow to close modal immediately after successful `jimakuDownloadFile` result. Added regression test (`src/renderer/modals/jimaku.test.ts`) that reproduces keyboard file-selection success path and asserts modal close state + `notifyOverlayModalClosed('jimaku')` emission. Kept failure path unchanged.

Also wired new test into `test:core:src` and `test:core:dist` package scripts.

<!-- SECTION:FINAL_SUMMARY:END -->
