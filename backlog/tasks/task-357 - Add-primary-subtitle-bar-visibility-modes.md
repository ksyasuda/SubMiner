---
id: TASK-357
title: Add primary subtitle bar visibility modes
status: Done
assignee:
  - Codex
created_date: '2026-05-13 03:17'
updated_date: '2026-05-13 03:24'
labels:
  - overlay
  - subtitles
  - config
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Update the primary subtitle bar visibility control so the `v` key cycles through the same mode model used by secondary subtitles: hidden, visible, and hover/auto. The primary and secondary subtitle bars must keep separate visibility state and config defaults, and primary visibility changes must identify the primary bar in OSD feedback.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Pressing `v` cycles primary subtitle bar visibility through hidden, visible, and hover/auto without changing secondary subtitle visibility.
- [x] #2 In hover/auto mode, the primary subtitle bar is normally hidden but becomes visible and interactable when hovering over its reserved location.
- [x] #3 Primary and secondary subtitle visibility modes have independent runtime state and config defaults.
- [x] #4 `subtitleStyle` exposes a primary subtitle visibility default that defaults to visible and validates invalid values with a warning/fallback.
- [x] #5 OSD feedback is shown when primary visibility mode changes and clearly identifies the primary subtitle bar.
- [x] #6 Relevant tests and config example/docs are updated.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing tests for primary subtitle mode config default/validation and renderer `v` cycling/OSD behavior.
2. Add primary subtitle mode type/config default under `subtitleStyle`, parse it with warning fallback, include it in generated example config.
3. Carry primary mode in renderer state/startup/hot-reload payloads independently from secondary mode.
4. Replace primary boolean toggle with `hidden -> visible -> hover` cycle and OSD text `Primary subtitle: <mode>`.
5. Add primary hover CSS mirroring secondary hover behavior so hover mode reserves a hit area and becomes interactable on hover.
6. Run focused tests, then broader relevant checks if time/environment allows.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented primary subtitle visibility as independent `PrimarySubMode` (`hidden | visible | hover`) with renderer state, startup default from `subtitleStyle.primaryDefaultMode`, hot-reload payload propagation, and primary-specific OSD. Focused tests were added before implementation and all focused tests passed. Full relevant gate passed: `bun run typecheck`, `bun run test:config`, `bun run test:fast`, `bun run test:env`, `bun run build`, `bun run test:smoke:dist`, `bun run docs:test`, `bun run docs:build`, `bun run format:check:src`, and `bun run changelog:lint`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Added `subtitleStyle.primaryDefaultMode` with default `visible`, enum validation, generated config examples, and renderer/hot-reload payload wiring.
- Changed the primary subtitle bar `v` action from a boolean hide/show toggle to independent `hidden | visible | hover` mode cycling with OSD text that identifies the primary subtitle.
- Added hover-mode CSS so the primary bar stays invisible until its location is hovered, then becomes visible and interactable.

Tests:
- `bun test src/config/resolve/subtitle-style.test.ts src/config/config.test.ts src/renderer/handlers/keyboard.test.ts src/renderer/subtitle-render.test.ts src/main/runtime/config-hot-reload-handlers.test.ts`
- `bun run typecheck`
- `bun run test:config`
- `bun run test:fast`
- `bun run test:env`
- `bun run build`
- `bun run test:smoke:dist`
- `bun run docs:test`
- `bun run docs:build`
- `bun run format:check:src`
- `bun run changelog:lint`
<!-- SECTION:FINAL_SUMMARY:END -->
