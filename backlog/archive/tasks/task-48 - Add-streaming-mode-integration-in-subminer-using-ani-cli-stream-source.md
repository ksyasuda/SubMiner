---
id: TASK-48
title: Add streaming mode integration in subminer using ani-cli stream source
status: To Do
assignee: []
created_date: '2026-02-14 06:01'
updated_date: '2026-02-14 08:19'
labels:
  - stream
  - ani-cli
  - jimaku
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Implement a new streaming mode so SubMiner can resolve and play episodes via ani-cli stream sources instead of existing file/download flow. The mode is enabled with a CLI flag (`-s` / `--stream`) and, when active, should prefer streamed playback and subtitle handling that keeps Japanese (or configured primary) subtitles as the main track. If a stream lacks Japanese subtitle tracks, fetch and inject subtitles from Jimaku and set them as primary subtitles according to SubMiner config.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Add a command-line option `-s`/`--stream` that enables streaming mode and is documented in help/config UX.
- [ ] #2 When streaming mode is enabled, resolve episode/video URLs using existing ani-cli stream selection logic (ported into SubMiner) and route playback to the resolved stream source.
- [ ] #3 If stream metadata contains a Japanese subtitle track, preserve that track as the primary subtitle stream path per current primary-subtitle selection behavior.
- [ ] #4 If no Japanese subtitle track is present in the stream metadata, fetch matching subtitles from Jimaku and load them into playback.
- [ ] #5 When loading Jimaku subtitles, replace/overwrite existing non-Japanese primary subtitle track behavior so the language configured as primary in SubMiner config is used instead.
- [ ] #6 Non-streaming mode continues to follow current behavior when `-s`/`--stream` is not set.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Superseded by TASK-51. Streaming via ani-cli in subminer was removed due Cloudflare 403/unreliable source metadata; re-evaluate later if reintroduced behind a feature flag and redesigned resolver/metadata pipeline.

<!-- SECTION:NOTES:END -->

## Definition of Done

<!-- DOD:BEGIN -->

- [ ] #1 CLI accepts both `-s` and `--stream` and enables streaming-specific behavior.
- [ ] #2 Streaming mode resolves streams through migrated ani-cli logic.
- [ ] #3 Japanese subs are preferred from stream metadata when available; Jimaku fallback is used only when absent.
- [ ] #4 Primary subtitle language in config determines which language is treated as default stream subtitle track after fallback.
- [ ] #5 Behavior is verified via test or manual checklist and documented in task notes.
<!-- DOD:END -->
