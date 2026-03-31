---
id: TASK-260
title: >-
  Fix macOS overlay subtitle sidebar passthrough without requiring a subtitle
  hover cycle
status: To Do
assignee: []
created_date: '2026-03-31 00:58'
labels:
  - bug
  - macos
  - overlay
  - subtitle-sidebar
  - passthrough
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/renderer/modals/subtitle-sidebar.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/renderer/overlay-mouse-ignore.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/renderer/handlers/mouse.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/main/overlay-runtime.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/overlay-visibility.ts
documentation:
  - docs/workflow/verification.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
On macOS, opening the overlay-layout subtitle sidebar should allow click-through outside the sidebar immediately. Users should not need to first hover subtitle content before passthrough/click-through starts working, including when no subtitle line is currently visible.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 With the overlay-layout subtitle sidebar open on macOS, areas outside the sidebar pass clicks through immediately after open without requiring a prior subtitle hover.
- [ ] #2 When no subtitle line is currently visible, opening the subtitle sidebar still leaves non-sidebar overlay regions click-through on macOS.
- [ ] #3 Regression coverage exercises the first-open/idle passthrough path so overlay interactivity does not depend on a later hover cycle.
<!-- AC:END -->
