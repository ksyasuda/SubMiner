---
id: TASK-38
title: Add user-friendly config validation errors on startup
status: To Do
assignee: []
created_date: '2026-02-14 02:02'
labels:
  - config
  - developer-experience
  - error-handling
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Improve config validation to surface clear, actionable error messages at startup when the user's config file has invalid values, missing required fields, or type mismatches.

## Motivation
The project has a config schema with validation, but invalid configs (wrong types, unknown keys after an upgrade, deprecated fields) can cause cryptic failures deep in service initialization rather than being caught and reported clearly at launch.

## Scope
1. Validate the full config against the schema at startup, before any services initialize
2. Collect all validation errors (don't fail on the first one) and present them as a summary
3. Show the specific field path, expected type, and actual value for each error
4. For deprecated or renamed fields, suggest the correct field name
5. Optionally show validation errors in a startup dialog (Electron dialog) rather than just console output
6. Allow the app to start with defaults for non-critical invalid fields, warning the user about the fallback

## Design constraints
- Must not block startup for non-critical warnings (e.g., unknown extra keys)
- Critical errors (e.g., invalid Anki field mappings) should prevent startup with a clear message
- Config file location should be shown in error output so users know what to edit
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 All config fields are validated against the schema before service initialization.
- [ ] #2 Validation errors show field path, expected type, and actual value.
- [ ] #3 Multiple errors are collected and shown together, not one at a time.
- [ ] #4 Deprecated or renamed fields produce a helpful migration suggestion.
- [ ] #5 Non-critical validation issues allow startup with defaults and a visible warning.
- [ ] #6 Critical validation failures prevent startup with a clear dialog or console message.
- [ ] #7 Config file path is shown in error output.
<!-- AC:END -->
