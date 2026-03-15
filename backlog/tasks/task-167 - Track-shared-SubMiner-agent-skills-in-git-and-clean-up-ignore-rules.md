---
id: TASK-167
title: Track shared SubMiner agent skills in git and clean up ignore rules
status: Done
assignee: []
created_date: '2026-03-13 05:46'
updated_date: '2026-03-13 05:47'
labels:
  - git
  - agents
  - repo-hygiene
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/.gitignore
  - >-
    /home/sudacode/projects/japanese/SubMiner/.agents/skills/subminer-change-verification/SKILL.md
  - >-
    /home/sudacode/projects/japanese/SubMiner/.agents/skills/subminer-scrum-master/SKILL.md
documentation:
  - /home/sudacode/projects/japanese/SubMiner/testing-plan.md
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Adjust the repository ignore rules so the shared SubMiner agent skill files can be committed while keeping unrelated local agent state ignored. Also ensure generated local verification artifacts like `.tmp/` do not pollute git status.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Root ignore rules allow the shared SubMiner skill files under `.agents/skills/` to be tracked without broadly unignoring local agent state
- [x] #2 The changed shared skill files appear in git status as trackable files after the ignore update
- [x] #3 Local generated verification artifact directories remain ignored so git status stays clean
- [x] #4 The updated ignore rules are minimal and scoped to the repo-shared skill files
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Updated `.gitignore` to keep `.agents` ignored by default while narrowly unignoring the repo-shared SubMiner skill files and verifier scripts.

Added `.tmp/` to the root ignore rules so local verification artifacts stop polluting `git status`.

Verified the result with `git status --untracked-files=all` and `git check-ignore -v`, confirming the shared skill files are now trackable and `.tmp/` remains ignored.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Adjusted the root `.gitignore` so the shared SubMiner agent skill files can be committed cleanly without broadly unignoring local agent state. The repo now tracks the shared `subminer-change-verification` skill files and the `subminer-scrum-master` skill doc, while `.tmp/` is ignored so generated verification artifacts do not pollute git status. Verified with `git status --untracked-files=all` and `git check-ignore -v` that the intended skill files are commit-ready and `.tmp/` remains ignored.
<!-- SECTION:FINAL_SUMMARY:END -->
