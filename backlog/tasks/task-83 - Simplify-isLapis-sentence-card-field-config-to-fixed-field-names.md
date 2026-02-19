---
id: TASK-83
title: Simplify isLapis sentence card field config to fixed field names
status: Done
assignee:
  - codex-main
created_date: '2026-02-19 08:38'
updated_date: '2026-02-19 08:40'
labels:
  - config
  - anki
  - cleanup
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Remove `sentenceCardSentenceField` and `sentenceCardAudioField` from the `ankiConnect.isLapis` config surface so Lapis mode always uses fixed field names `Sentence` and `SentenceAudio`. Update types/defaults/docs/examples and adjust integration code to stop reading those keys.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `isLapis` config no longer exposes `sentenceCardSentenceField` or `sentenceCardAudioField` in types/defaults/examples/docs
- [x] #2 Anki integration uses fixed sentence-card field names `Sentence` and `SentenceAudio` when Lapis/Kiku sentence-card flow runs
- [x] #3 Project build and relevant tests pass after removal
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Removed `sentenceCardSentenceField` and `sentenceCardAudioField` from `isLapis` config types/defaults (`src/types.ts`, `src/config/definitions.ts`).

Updated Anki integration to always use fixed sentence-card field names `Sentence` and `SentenceAudio` in effective sentence-card config (`src/anki-integration.ts`).

Hardened config resolution for `ankiConnect.isLapis` to only read `enabled` and `sentenceCardModel`, and emit warnings when deprecated sentence/audio field override keys are provided (`src/config/service.ts`).

Updated examples/docs to remove the two keys and document fixed mapping (`config.example.jsonc`, `docs/public/config.example.jsonc`, `docs/configuration.md`, `docs/anki-integration.md`, `docs/mining-workflow.md`).

Added regression coverage `ignores deprecated isLapis sentence-card field overrides` in `src/config/config.test.ts`.

Verification: `bun run build && bun run test:config:dist` passed (33/33).
<!-- SECTION:NOTES:END -->
