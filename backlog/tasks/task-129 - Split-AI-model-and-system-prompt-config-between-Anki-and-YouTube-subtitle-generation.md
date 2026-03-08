---
id: TASK-129
title: >-
  Split AI model and system prompt config between Anki and YouTube subtitle
  generation
status: Done
assignee: []
created_date: '2026-03-08 09:40'
updated_date: '2026-03-08 09:57'
labels:
  - config
  - ai
  - anki
  - youtube-subgen
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
The current top-level shared AI config forces Anki translation and YouTube subtitle fixing to share the same model and system prompt, which caused subtitle-fix requests to inherit a translation prompt and translate Japanese primary subtitles into English. Refactor config so provider credentials stay shared while model and system prompt can be configured per feature.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Anki integration can use its own AI model and system prompt independently of YouTube subtitle generation.
- [x] #2 YouTube subtitle generation can use its own AI model and system prompt independently of Anki integration.
- [x] #3 Existing shared provider credentials remain reusable without duplicating API key/base URL config.
- [x] #4 Config example, defaults, validation, and regression tests cover the new per-feature override shape.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added per-feature AI model/systemPrompt overrides for Anki and YouTube subtitle generation while keeping shared provider transport settings reusable. Anki now accepts `ankiConnect.ai` object config with `enabled`, `model`, and `systemPrompt`; YouTube subtitle generation accepts `youtubeSubgen.ai` overrides and merges them over the shared AI provider config. Updated config resolution, launcher parsing, runtime wiring, hot-reload handling, example config, and regression coverage.
<!-- SECTION:FINAL_SUMMARY:END -->
