---
id: TASK-300
title: Fix transparent subtitle hover background config
status: Done
assignee: []
created_date: '2026-04-26 03:23'
updated_date: '2026-04-26 03:26'
labels:
  - bug
  - overlay
  - config
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
User reports setting subtitleStyle.hoverTokenBackgroundColor to transparent still renders default hover background in overlay subtitles.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Transparent hoverTokenBackgroundColor is accepted by config resolution.
- [x] #2 Renderer applies transparent hover token background instead of falling back to default.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Reproduce config alias behavior with a failing config test.
2. Map subtitleStyle.hoverBackground to hoverTokenBackgroundColor in config resolution while keeping canonical key precedence.
3. Add renderer regression for transparent hover token background CSS variable.
4. Update docs and changelog fragment; run focused verification.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Local user config used subtitleStyle.hoverBackground, which was ignored because only subtitleStyle.hoverTokenBackgroundColor was recognized. Canonical key still takes precedence when both are present.

Verification passed: bun run test:config:src; bun test src/renderer/subtitle-render.test.ts; bun run changelog:lint; bun run docs:test; bun run docs:build.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented config compatibility for transparent hover token backgrounds. `subtitleStyle.hoverBackground` now maps to the canonical `subtitleStyle.hoverTokenBackgroundColor` during resolution, preserving canonical key precedence. Added regression coverage for the alias and renderer handling of `transparent`, documented the alias, and added a changelog fragment.

Verification: `bun run test:config:src`; `bun test src/renderer/subtitle-render.test.ts`; `bun run changelog:lint`; `bun run docs:test`; `bun run docs:build`.
<!-- SECTION:FINAL_SUMMARY:END -->
