---
id: TASK-121
title: >-
  Fix YouTube manual subtitle selection regression when downloadable tracks
  exist
status: Done
assignee:
  - '@codex'
created_date: '2026-03-08 05:37'
updated_date: '2026-03-16 05:13'
labels:
  - bug
  - youtube
  - subtitles
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/launcher/youtube/manual-subs.ts
  - /Users/sudacode/projects/japanese/SubMiner/launcher/youtube/orchestrator.ts
  - 'https://www.youtube.com/watch?v=MXzQRLmN9hE'
priority: high
ordinal: 64500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Ensure launcher YouTube subtitle generation reuses downloadable manual subtitle tracks when the video already has requested languages available, instead of falling back to whisper generation. Reproduce against videos like MXzQRLmN9hE that expose manual en/ja subtitles via yt-dlp.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 When requested primary/secondary manual YouTube subtitle tracks exist, planning selects them and schedules no whisper generation for those tracks.
- [x] #2 Filename normalization handles manual subtitle outputs produced by yt-dlp for language-tagged downloads.
- [x] #3 Automated tests cover the reproduced manual en/ja selection case.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Reproduced against https://www.youtube.com/watch?v=MXzQRLmN9hE with yt-dlp --list-subs: manual zh/en/ja/ko subtitle tracks are available from YouTube.

Adjusted launcher YouTube orchestration so detected manual subtitle tracks suppress whisper generation but are no longer materialized as external subtitle files. SubMiner now relies on the native YouTube/mpv subtitle tracks for those languages.

Added orchestration tests covering the manual-track reuse plan and ran a direct runtime probe against MXzQRLmN9hE. Probe result: primary/secondary native tracks detected, no external subtitle aliases emitted, output directory remained empty.

Verification: bun test launcher/youtube/orchestrator.test.ts launcher/config-domain-parsers.test.ts launcher/mpv.test.ts passed; bun run typecheck passed.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed the YouTube subtitle regression where videos with real downloadable subtitle tracks still ended up with duplicate external subtitle files. Manual subtitle availability now suppresses whisper generation and external subtitle publication, so videos like MXzQRLmN9hE use the native YouTube/mpv subtitle tracks directly. Launcher preprocess logging was also updated to report native subtitle availability instead of misleading missing statuses.
<!-- SECTION:FINAL_SUMMARY:END -->
