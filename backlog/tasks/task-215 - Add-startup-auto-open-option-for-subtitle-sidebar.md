---
id: TASK-215
title: Add startup auto-open option for subtitle sidebar
status: In Progress
assignee: []
created_date: '2026-03-21 11:35'
updated_date: '2026-03-21 11:35'
labels:
  - feature
  - ux
  - overlay
  - subtitles
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/src/types.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/config/definitions/defaults-subtitle.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/config/resolve/subtitle-domains.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/renderer/modals/subtitle-sidebar.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/renderer/renderer.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add a subtitle sidebar config option that auto-opens the sidebar once during overlay startup. The option should default to `false`, only apply when the sidebar feature is enabled, and should not force the sidebar back open later in the same session after manual close or later visibility changes.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `subtitleSidebar.autoOpen` is available in config with default `false`.
- [x] #2 When enabled, overlay startup opens the subtitle sidebar once after initial sidebar config/snapshot load.
- [x] #3 Regression coverage covers config resolution and startup-only auto-open behavior.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-03-21: Added `subtitleSidebar.autoOpen` to types/defaults/config registry and resolver. Renderer bootstrap now calls a startup-only subtitle sidebar helper after the initial snapshot refresh. Modal regression coverage verifies startup auto-open requires both `enabled` and `autoOpen`.
<!-- SECTION:NOTES:END -->
