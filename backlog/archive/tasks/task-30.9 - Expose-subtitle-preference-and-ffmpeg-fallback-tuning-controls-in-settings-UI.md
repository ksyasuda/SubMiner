---
id: TASK-30.9
title: Expose subtitle preference and ffmpeg fallback tuning controls in settings UI
status: To Do
assignee: []
created_date: '2026-02-13 18:42'
labels: []
dependencies:
  - TASK-30.7
  - TASK-30.8
parent_task_id: TASK-30
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Add user-configurable controls for Aniyomi streaming subtitle behavior, including preferred language profile, soft-vs-hard source preference, ffmpeg-assisted hard-sub removal behavior, and policy toggles so quality and fallback behavior can be tuned without code changes.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Add settings UI fields to define preferred subtitle/audiotrack language order (e.g., en, ja) and enable/disable hard-sub fallback mode.
- [ ] #2 Add explicit toggle for enabling hard-sub stripping via ffmpeg and configurable timeout/quality limits to avoid long waits.
- [ ] #3 Expose source ranking preferences for soft-sub vs hard-sub sources and optional fallback to native/transcoded source when preferred modes are unavailable.
- [ ] #4 Persist settings in existing config schema with migration-safe defaults and include clear validation for invalid/unsupported values.
- [ ] #5 Show current effective policy in streaming UI (for debugging): source selected + reason + whether subtitles are soft/hard and if ffmpeg path is active.
- [ ] #6 Add user-facing explanatory text and warnings for ffmpeg dependency, expected CPU/cpu cost, and compatibility limits.
- [ ] #7 Log and display when user-adjusted policy changes alter a previously preferred source choice during runtime.
<!-- AC:END -->
