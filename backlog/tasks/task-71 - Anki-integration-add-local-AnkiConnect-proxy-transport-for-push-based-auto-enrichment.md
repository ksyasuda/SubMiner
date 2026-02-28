---
id: TASK-71
title: >-
  Anki integration: add local AnkiConnect proxy transport for push-based
  auto-enrichment
status: Done
assignee: []
created_date: '2026-02-28 02:38'
labels: []
dependencies: []
references:
  - src/anki-integration/anki-connect-proxy.ts
  - src/anki-integration/anki-connect-proxy.test.ts
  - src/anki-integration.ts
  - src/config/resolve/anki-connect.ts
  - src/core/services/tokenizer/yomitan-parser-runtime.ts
  - src/core/services/tokenizer/yomitan-parser-runtime.test.ts
  - docs/anki-integration.md
  - config.example.jsonc
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Scope: Current unmerged working-tree changes implement an optional local AnkiConnect-compatible proxy and transport switching for card enrichment.

Delivered behavior:
- Added proxy server that forwards AnkiConnect requests and enqueues addNote/addNotes note IDs for post-create enrichment, with de-duplication and loop-configuration protection.
- Added config schema/defaults/resolution for ankiConnect.proxy (enabled, host, port, upstreamUrl) with validation warnings and fallback behavior.
- Runtime now supports transport switching (polling vs proxy) and restarts transport when runtime config patches change transport keys.
- Added Yomitan default-profile server sync helper to keep bundled parser profile aligned with configured Anki endpoint.
- Updated user docs/config examples for proxy mode setup, troubleshooting, and mining workflow behavior.

Risk/impact context:
- New network surface on local host/port; correctness depends on safe proxy upstream configuration and robust response handling.
- Tests added for proxy queue behavior, config resolution, and parser sync routines.
<!-- SECTION:DESCRIPTION:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Completed implementation in branch working tree; ready to merge once local changes are committed and test gate passes.
<!-- SECTION:FINAL_SUMMARY:END -->
