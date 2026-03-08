---
id: TASK-117
title: >-
  Replace YouTube subtitle generation with pure TypeScript pipeline and shared
  AI config
status: Done
assignee:
  - codex
created_date: '2026-03-08 03:16'
updated_date: '2026-03-08 03:35'
labels: []
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/launcher/youtube.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/anki-integration/ai.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/types.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/config/definitions/defaults-integrations.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/config/resolve/subtitle-domains.ts
  - /Users/sudacode/projects/japanese/SubMiner/config.example.jsonc
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Replace the launcher YouTube subtitle generation flow with a pure TypeScript pipeline that prefers real downloadable YouTube subtitles, never uses YouTube auto-generated subtitles, locally generates missing tracks with whisper.cpp, and can optionally fix generated subtitles via a shared OpenAI-compatible AI provider config. This feature also introduces a breaking config cleanup: move provider settings to a new top-level ai section and reduce ankiConnect.ai to a boolean feature toggle.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launcher YouTube subtitle generation prefers downloadable manual YouTube subtitles, never uses YouTube auto-generated subtitles, and locally generates only missing tracks with whisper.cpp.
- [x] #2 Generated whisper subtitle tracks can optionally be post-processed with an OpenAI-compatible AI provider using shared top-level ai config, with validation and fallback to raw whisper output on failure.
- [x] #3 Configuration is updated so top-level ai is canonical shared provider config, ankiConnect.ai is boolean-only, and youtubeSubgen includes whisperVadModel, whisperThreads, and fixWithAi.
- [x] #4 Launcher CLI/config parsing, config example, and docs reflect the new breaking config shape with no migration layer.
- [x] #5 Automated tests cover the new YouTube generation behavior, AI-fix fallback/validation behavior, shared AI config usage, and breaking config validation.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Introduce canonical top-level ai config plus youtubeSubgen runtime knobs (whisperVadModel, whisperThreads, fixWithAi) and convert ankiConnect.ai to a boolean-only toggle across types, defaults, validation, option registries, launcher config parsing, and config example/docs.
2. Extract shared OpenAI-compatible AI client helpers from the current Anki translation code, including base URL normalization, API key / apiKeyCommand resolution, timeout handling, and response text extraction.
3. Update Anki translation flow and hot-reload/runtime plumbing to consume global ai config while treating ankiConnect.ai as a feature gate only.
4. Replace launcher/youtube.ts with a modular launcher/youtube pipeline that fetches only manual YouTube subtitles, generates missing tracks locally with ffmpeg + whisper.cpp + optional VAD/thread controls, and preserves preprocess/automatic playback behavior.
5. Add optional AI subtitle-fix processing for whisper-generated tracks using the shared ai client, with strict SRT batching/validation and fallback to raw whisper output on provider or format failure.
6. Expand automated coverage for config validation, shared AI usage, launcher config parsing, and YouTube subtitle generation behavior including removal of yt-dlp auto-subs and AI-fix fallback rules.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented pure TypeScript launcher/youtube pipeline modules for manual subtitle fetch, audio extraction, whisper runs, SRT utilities, and optional AI subtitle fixing. Removed yt-dlp auto-subtitle usage from the generation path.

Added shared top-level ai config plus shared AI client helpers; converted ankiConnect.ai to a boolean feature gate and updated Anki runtime wiring to consume global ai config.

Updated launcher config parsing, config template sections, and config.example.jsonc for the breaking config shape including youtubeSubgen.whisperVadModel, youtubeSubgen.whisperThreads, and youtubeSubgen.fixWithAi.

Verification: bun run test:config:src passed; targeted AI/Anki/runtime tests passed; bun run typecheck passed. bun run test:launcher:unit:src reported one unrelated existing failure in launcher/aniskip-metadata.test.ts (resolveAniSkipMetadataForFile resolves MAL id and intro payload).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Replaced the launcher YouTube subtitle flow with a modular TypeScript pipeline that prefers manual YouTube subtitles, transcribes only missing tracks with whisper.cpp, and can optionally post-fix whisper output through a shared OpenAI-compatible AI client with strict SRT validation/fallback. Introduced canonical top-level ai config, reduced ankiConnect.ai to a boolean feature gate, updated launcher/config parsing and checked-in config artifacts, and added coverage for YouTube orchestration, whisper args, SRT validation, AI fix behavior, and breaking config validation.
<!-- SECTION:FINAL_SUMMARY:END -->
