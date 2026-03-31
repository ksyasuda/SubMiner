---
id: TASK-194
title: App-owned YouTube subtitle picker flow
status: Done
assignee: []
created_date: '2026-03-18 07:52'
updated_date: '2026-03-23 03:22'
labels: []
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/launcher/youtube/orchestrator.ts
  - /home/sudacode/projects/japanese/SubMiner/launcher/youtube/manual-subs.ts
  - /home/sudacode/projects/japanese/SubMiner/src/core/services/tokenizer.ts
documentation:
  - /home/sudacode/projects/japanese/SubMiner/youtube.md
priority: medium
ordinal: 137500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Replace the YouTube subtitle-generation-first flow with an app-owned picker flow that boots mpv paused, opens an overlay track picker, downloads selected subtitles into external subtitle files, and preserves generation as an explicit mode. Keep the existing SubMiner tokenization and annotation pipeline as the downstream consumer of downloaded subtitle files.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launcher and app expose YouTube subtitle acquisition modes `download` and `generate`, with `download` as the default.
- [x] #2 YouTube playback boots mpv paused and presents an overlay selection UI for primary and secondary subtitle choices.
- [x] #3 Selected YouTube subtitle tracks are downloaded to external subtitle files and loaded into mpv before playback resumes.
- [x] #4 `generate` mode preserves the existing subtitle generation path as an explicit opt-in behavior.
- [x] #5 Downloaded YouTube subtitle files integrate with the existing SubMiner subtitle/tokenization/annotation pipeline without regressing current overlay behavior.
- [x] #6 Tests cover mode selection, subtitle-track enumeration/selection flow, and the paused bootstrap plus app handoff path.
- [x] #7 User-facing config and launcher docs are updated to describe the new modes and default behavior.
<!-- AC:END -->
