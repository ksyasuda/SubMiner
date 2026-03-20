---
id: TASK-100
title: Add configurable texthooker startup launch
status: Done
assignee: []
created_date: '2026-03-06 23:30'
updated_date: '2026-03-16 05:13'
labels: []
dependencies: []
priority: medium
ordinal: 11010
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add a config option under `texthooker` to launch the built-in texthooker server automatically when SubMiner starts.

Scope:

- Add `texthooker.launchAtStartup`.
- Default to `true`.
- Start the existing texthooker server during normal app startup when enabled.
- Keep `texthooker.openBrowser` as separate behavior.
- Add regression coverage and update generated config docs/example.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Default config enables automatic texthooker startup.
- [x] #2 Config parser accepts valid boolean values and warns on invalid values.
- [x] #3 App-ready startup launches texthooker when enabled.
- [x] #4 Generated config template/example documents the new option.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added `texthooker.launchAtStartup` with a default of `true`, wired it through config defaults/validation/template generation, and started the existing texthooker server during app-ready startup without coupling it to browser auto-open behavior.

Also added regression coverage for config parsing/template output and app-ready dependency wiring, then regenerated the checked-in config example artifacts.
<!-- SECTION:FINAL_SUMMARY:END -->
