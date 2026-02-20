---
id: TASK-91
title: Add config toggle to preserve visible overlay subtitle line breaks
status: Done
assignee: []
created_date: '2026-02-20 06:35'
updated_date: '2026-02-20 06:42'
labels: []
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add a `subtitleStyle` config option that keeps visible-overlay subtitle line breaks (newline/carriage-return normalized to line breaks) instead of flattening them to spaces. Default should preserve current behavior for consistency with texthooker.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 New config option exists with default disabled and validation/docs coverage.
- [x] #2 When enabled, visible overlay preserves subtitle line breaks while rendering tokenized subtitles.
- [x] #3 When disabled, current rendering behavior remains unchanged.
- [x] #4 Relevant config + renderer tests pass.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added `subtitleStyle.preserveLineBreaks` (default `false`) to types/default config/registry/config validation and docs/example configs.

Renderer now supports line-break-preserving token output via `alignTokensToSourceText` in `src/renderer/subtitle-render.ts`, which inserts source-text separators (including `\n`) between token spans when enabled.

Validation:
- `bun run build && node --test dist/config/config.test.js dist/renderer/subtitle-render.test.js`
<!-- SECTION:NOTES:END -->
