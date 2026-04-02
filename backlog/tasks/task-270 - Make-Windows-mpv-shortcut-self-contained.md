---
id: TASK-270
title: Make Windows mpv shortcut self-contained
status: Done
assignee: []
created_date: '2026-04-02 07:13'
updated_date: '2026-04-02 07:19'
labels: []
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Remove the Windows mpv shortcut's dependency on a pre-existing mpv profile so the installer-created `SubMiner mpv` flow works out of the box without requiring the user to edit `mpv.conf`. Keep docs aligned with the new behavior and preserve the optional profile guidance for manual mpv usage.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `SubMiner.exe --launch-mpv` launches mpv with SubMiner's required default args without requiring an mpv profile named `subminer`.
- [x] #2 Windows shortcut/help/docs no longer describe `--launch-mpv` as depending on the SubMiner mpv profile.
- [x] #3 Automated tests cover the Windows launch args behavior and pass after the change.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Updated the Windows `--launch-mpv` path to pass SubMiner's default mpv args directly instead of requiring `--profile=subminer`. Adjusted Windows shortcut/help text to describe the self-contained defaults-based launch, and updated Windows docs to state that `mpv.conf` is not required for the shortcut path while preserving the optional profile guidance for manual mpv launches.
<!-- SECTION:FINAL_SUMMARY:END -->
