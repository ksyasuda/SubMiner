---
id: TASK-72
title: Make startup config loading strict with clear user-facing errors
status: To Do
assignee: []
created_date: '2026-02-18 11:35'
updated_date: '2026-02-18 11:35'
labels:
  - config
  - ux
  - reliability
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Startup config loading currently falls back to default config when strict parse fails, while hot-reload path reports strict errors. This mismatch can hide broken config at startup and cause surprising runtime behavior.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Align startup loading semantics with strict reload behavior.
2. Surface parse/validation errors to user via clear logs/OSD/notification where appropriate.
3. Define fallback policy explicitly (fail fast vs safe mode) and apply consistently.
4. Add regression tests for malformed JSON/JSONC at startup.
5. Document startup error behavior in troubleshooting/config docs.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Startup and hot-reload use consistent strictness semantics
- [ ] #2 Malformed config surfaces explicit actionable error
- [ ] #3 Tests cover malformed startup config path
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Config and startup tests pass
- [ ] #2 Docs updated for new startup error behavior
<!-- DOD:END -->

