---
id: TASK-48.6
title: Wire -s/--stream mode to Ani-cli and Jimaku subtitle fallback
status: To Do
assignee: []
created_date: '2026-02-14 06:06'
updated_date: '2026-02-14 08:19'
labels: []
dependencies: []
references:
  - ani-cli/ani-cli
  - src/jimaku/utils.ts
  - src/core/services/anki-jimaku-service.ts
documentation:
  - ani-cli/README.md
  - subminer
  - src/cli/help.ts
parent_task_id: TASK-48
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Implement SubMiner streaming mode end-to-end behind a `-s`/`--stream` flag. In stream mode, use the vendored ani-cli resolution flow to get playable stream URLs instead of local file/YouTube URL handling. If resolved streams do not expose Japanese subtitles, fetch matching subtitles from Jimaku and load them into mpv as the active primary subtitle track, overwriting the current non-primary/non-Japanese default according to subminer primary-subtitle configuration.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 When `subminer -s` is used, resolution should pass a search/query through ani-cli stream logic and play the resolved stream source.
- [ ] #2 If the stream includes a Japanese subtitle track, preserve and select the configured primary subtitle language behavior without Jimaku injection.
- [ ] #3 If no Japanese (or configured primary language) subtitle exists in stream metadata, fetch and inject Jimaku subtitles before playback starts.
- [ ] #4 Loaded Jimaku subtitles should become the selected primary subtitle track, replacing any existing default non-primary subtitle in that context.
- [ ] #5 When `-s` is not passed, non-streaming behavior remains unchanged.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Superseded by TASK-51. End-to-end stream wiring to ani-cli is deferred.
<!-- SECTION:NOTES:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 CLI exposes both `-s` and `--stream` in help/config and validation.
- [ ] #2 Implementation includes a clear fallback path when stream subtitles are absent and Jimaku search/download fails gracefully.
- [ ] #3 Subtitles loading path avoids temp-file leaks; temporary media/subtitle artifacts are cleaned up on exit.
- [ ] #4 At least one verification step (manual or test) confirms stream mode path works for an episode with and without Japanese stream subtitles.
<!-- DOD:END -->
