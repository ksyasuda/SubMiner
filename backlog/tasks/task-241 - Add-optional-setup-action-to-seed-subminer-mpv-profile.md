id: TASK-241
title: Add optional setup action to seed SubMiner mpv profile
type: feature
status: Open
assignee: []
created_date: '2026-03-27 11:22'
updated_date: '2026-03-27 11:22'
labels:
  - setup
  - mpv
  - docs
  - ux
dependencies: []
references: []
documentation:
  - /home/sudacode/projects/japanese/SubMiner/docs-site/usage.md
  - /home/sudacode/projects/japanese/SubMiner/docs-site/launcher-script.md
ordinal: 24100
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add an optional control in the first-run / setup flow to write or update the user’s mpv configuration with SubMiner-recommended defaults (especially the `subminer` profile), so users can recover from a missing profile without manual config editing.

The docs for launcher usage must explicitly state that SubMiner’s Windows mpv launcher path runs mpv with `--profile=subminer` by default.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->
- [ ] #1 Add an optional setup UI action/button to generate or overwrite a user-confirmed mpv config that includes a `subminer` profile.
- [ ] #2 The action should be non-destructive by default, show diff/contents before write, and support append/update mode when other mpv settings already exist.
- [ ] #3 Document how to resolve the missing-profile scenario and clearly state that the SubMiner mpv launcher runs with `--profile=subminer` by default (`--launch-mpv` / Windows mpv shortcut path).
- [ ] #4 Add/adjust setup validation messaging so users are not blocked if `subminer` profile is initially missing, but can opt into one-click setup recovery.
- [ ] #5 Include a short verification path for both Windows and non-Windows flows (for example dry-run + write path).
<!-- AC:END -->
