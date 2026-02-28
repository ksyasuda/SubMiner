---
id: TASK-72
title: 'macOS config validation UX: show full warning details in native dialog'
status: Done
assignee: []
created_date: '2026-02-28 02:38'
labels: []
dependencies: []
references:
  - 'commit:cc2f9ef'
  - src/main/config-validation.ts
  - src/main/runtime/startup-config.ts
  - docs/configuration.md
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Scope: Commit cc2f9ef improves startup config-warning visibility on macOS by ensuring full details are surfaced in the native UI path and reflected in docs.

Delivered behavior:
- Config validation/runtime wiring updated so macOS users can access complete warning details instead of truncated notification-only text.
- Added/updated tests around config validation and startup config warning flows.
- Updated configuration docs to clarify platform-specific warning presentation behavior.

Risk/impact context:
- Low runtime risk; primarily user-facing diagnostics clarity improvement.
<!-- SECTION:DESCRIPTION:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Completed small follow-up fix to reduce config-debug friction on macOS.
<!-- SECTION:FINAL_SUMMARY:END -->
