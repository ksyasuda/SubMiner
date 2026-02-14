---
id: TASK-30.7
title: >-
  Add English-source preference + hard-sub stripping workflow in Aniyomi
  streaming path
status: To Do
assignee: []
created_date: '2026-02-13 18:41'
labels: []
dependencies:
  - TASK-30.2
  - TASK-30.5
parent_task_id: TASK-30
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Improve the Aniyomi/anime extension streaming flow to prefer English-capable sources with soft subtitles, and automatically recover when only hard-subbed streams are available by stripping embedded subtitles with ffmpeg and attaching external Jimaku subtitle files into mpv.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 During source scoring/selection, prefer providers/sources that declare or expose soft subtitles for English audio or subtitle tracks over hard-subbed alternatives.
- [ ] #2 Add a config option for preferred language targets (default English) and fallback policy (favor soft subtitles, then hard-sub fallback).
- [ ] #3 Detect when a resolved stream is hard-sub-only and a soft-sub source is unavailable for the same episode.
- [ ] #4 When hard subs are used, attempt to generate a subtitle-less playback stream path using ffmpeg and feed an external subtitle file (including Jimaku-provided `.ass/.srt`) into the mpv playback path.
- [ ] #5 Preserve stream selection metadata (language tags, subtitle type, availability state) for UI decisions and error messaging.
- [ ] #6 If ffmpeg conversion is not possible/disabled or fails, surface a clear status that explains fallback behavior instead of silent failure.
- [ ] #7 Integrate subtitle source preferences with Jimaku so user-fetched or resolved subtitles are preferred for burn-in removal cases where supported.
- [ ] #8 Add handling for unsupported codecs/containers and provider limitations so direct passthrough is still used when hard-sub stripping is unsafe.
- [ ] #9 Document new behavior in feature docs/FAQ: how sources are ranked, what hard-sub stripping does, and known compatibility limitations.
<!-- AC:END -->
