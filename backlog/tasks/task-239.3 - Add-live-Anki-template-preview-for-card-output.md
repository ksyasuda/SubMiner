---
id: TASK-239.3
title: Add live Anki template preview for card output
status: To Do
assignee: []
created_date: '2026-03-26 20:49'
labels:
  - feature
  - anki
  - ux
milestone: m-2
dependencies: []
references:
  - src/anki-integration.ts
  - src/anki-integration/card-creation.ts
  - src/config/resolve/anki-connect.ts
  - src/renderer
parent_task_id: TASK-239
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Users currently have to infer what card output will look like from config fields and post-mine results. Add a live preview surface that shows the resolved card template output before mining so users can catch broken field mappings, missing media, or undesirable formatting earlier.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->
- [ ] #1 Users can open a preview that renders the resolved front/back field output for the current note/card template configuration.
- [ ] #2 The preview clearly surfaces missing or unmapped fields instead of silently showing blank content.
- [ ] #3 Preview generation uses the same transformation logic as the live card-creation path so it stays trustworthy.
- [ ] #4 The first slice works with representative sample mining payloads and handles missing optional media gracefully.
- [ ] #5 Tests cover preview rendering for at least one valid and one invalid/missing-field configuration.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Identify the current card-creation data path and extract any logic needed to render a preview without duplicating transformation rules.
2. Add a focused preview UI in the most relevant existing configuration/setup surface.
3. Surface validation/warning states for empty mappings, missing fields, and media-dependent outputs.
4. Verify with Anki integration tests plus renderer coverage for preview states.
<!-- SECTION:PLAN:END -->
