---
id: TASK-294
title: Fix annotated subtitle tokens to honor subtitleStyle
status: Done
assignee: []
created_date: '2026-04-25 23:04'
updated_date: '2026-04-25 23:07'
labels:
  - subtitles
  - renderer
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Annotated token spans should inherit the configured subtitleStyle typography and only use annotation metadata to change token color.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Tokenized/annotated subtitles preserve configured base subtitle typography such as font family, size, weight, line height, letter spacing, text rendering, and text shadow.
- [x] #2 Known/N+1/JLPT/frequency/name-match annotations affect token color only, plus existing token metadata/hover affordances.
- [x] #3 A renderer regression test covers annotated token rendering with custom subtitleStyle.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Updated renderer subtitle annotation CSS so known/N+1/name/JLPT/frequency classes no longer override typography with token-specific shadows, underlines, padding, or hover font-weight. Added regression coverage using the user's custom subtitleStyle shape to verify annotated token spans inherit base typography and annotation CSS changes token color only. Verified with `bun test src/renderer/subtitle-render.test.ts`, `bun run typecheck`, and `bun run test:fast`.
<!-- SECTION:FINAL_SUMMARY:END -->
