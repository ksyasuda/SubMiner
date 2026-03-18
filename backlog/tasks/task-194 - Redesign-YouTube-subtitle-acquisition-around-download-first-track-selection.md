---
id: TASK-194
title: Redesign YouTube subtitle acquisition around download-first track selection
status: To Do
assignee: []
created_date: '2026-03-18 07:52'
labels: []
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/launcher/youtube/orchestrator.ts
  - /home/sudacode/projects/japanese/SubMiner/launcher/youtube/manual-subs.ts
  - /home/sudacode/projects/japanese/SubMiner/src/core/services/tokenizer.ts
documentation:
  - /home/sudacode/projects/japanese/SubMiner/youtube.md
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Replace the current YouTube subtitle-generation-first flow with a download-first flow that enumerates available YouTube subtitle tracks, prompts for primary and secondary track selection before playback, downloads selected tracks into external subtitle files for mpv, and preserves generation as an explicit mode and as fallback behavior in auto mode. Keep the existing SubMiner tokenization and annotation pipeline as the downstream consumer of downloaded subtitle files.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Launcher and config expose YouTube subtitle acquisition modes `download`, `generate`, and `auto`, with `download` as the default for launcher YouTube playback.
- [ ] #2 YouTube playback enumerates available subtitle tracks before mpv launch and presents a selection UI that supports primary and secondary subtitle choices.
- [ ] #3 Selected YouTube subtitle tracks are downloaded to external subtitle files and loaded into mpv before playback starts when download mode succeeds.
- [ ] #4 `auto` mode attempts download-first for the selected tracks and falls back to generation only when required tracks cannot be downloaded or download fails.
- [ ] #5 `generate` mode preserves the existing whisper/AI generation path as an explicit opt-in behavior.
- [ ] #6 Downloaded YouTube subtitle files integrate with the existing SubMiner subtitle/tokenization/annotation pipeline without regressing current overlay behavior.
- [ ] #7 Tests cover mode selection, subtitle-track enumeration/selection flow, download-first success path, and fallback behavior for auto mode.
- [ ] #8 User-facing config and launcher docs are updated to describe the new modes and default behavior.
<!-- AC:END -->
