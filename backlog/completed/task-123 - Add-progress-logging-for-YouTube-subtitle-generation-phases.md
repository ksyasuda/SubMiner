---
id: TASK-123
title: Add progress logging for YouTube subtitle generation phases
status: Done
assignee:
  - '@codex'
created_date: '2026-03-08 07:07'
updated_date: '2026-03-16 05:13'
labels:
  - ux
  - logging
  - youtube
  - subtitles
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/launcher/youtube/orchestrator.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/launcher/youtube/audio-extraction.ts
  - /Users/sudacode/projects/japanese/SubMiner/launcher/youtube/whisper.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/launcher/youtube/subtitle-fix-ai.ts
priority: medium
ordinal: 62500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Improve launcher YouTube subtitle generation observability so users can tell that work is happening and roughly how long each phase is taking. Cover manual subtitle probe, audio extraction, ffmpeg prep, whisper generation, and optional AI subtitle fix phases without flooding normal logs.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Users see clear info-level phase logs for YouTube subtitle generation work including subtitle probe, fallback audio extraction, whisper, and optional AI fix phases.
- [x] #2 Long-running phases surface elapsed-time progress or explicit start/finish timing so it is obvious the process is still active.
- [x] #3 Automated tests cover the new logging/progress helper behavior where practical.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented a shared timed YouTube phase logger in launcher/youtube/progress.ts with info-level start/finish messages and warn-level failure messages that include elapsed time.

Wired phase logging into YouTube metadata probe, manual subtitle probe, fallback audio extraction, ffmpeg whisper prep, whisper primary/secondary generation, and optional AI subtitle fix phases.

Verification: bun test launcher/youtube/progress.test.ts launcher/youtube/orchestrator.test.ts passed; bun run typecheck passed.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added clear phase-level observability for YouTube subtitle generation without noisy tool output. Users now see start/finish logs with elapsed time for subtitle probe, fallback audio extraction, ffmpeg prep, whisper generation, and optional AI subtitle-fix phases, making it obvious when generation is active and roughly how long each step took.
<!-- SECTION:FINAL_SUMMARY:END -->
