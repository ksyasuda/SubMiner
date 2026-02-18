---
id: TASK-66
title: Add AnkiConnect tagging for mined and updated cards
status: Done
assignee: []
created_date: '2026-02-18 09:23'
updated_date: '2026-02-18 09:24'
labels:
  - anki
  - config
  - enhancement
dependencies: []
references:
  - src/anki-connect.ts
  - src/anki-integration.ts
  - src/anki-integration/card-creation.ts
  - src/config/definitions.ts
  - src/config/service.ts
  - src/config/config.test.ts
  - docs/configuration.md
  - config.example.jsonc
  - docs/public/config.example.jsonc
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Support configuring tags applied to cards created or updated through SubMiner's AnkiConnect workflow. Default behavior should add the `SubMiner` tag to all mined/updated cards, while allowing users to override or disable automatic tagging via config.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 AnkiConnect note creation supports passing tags.
- [x] #2 Existing-card update flows add configured tags via AnkiConnect after updates.
- [x] #3 Default config includes `ankiConnect.tags: ["SubMiner"]`.
- [x] #4 Config parsing validates `ankiConnect.tags` and falls back safely on invalid values.
- [x] #5 User-facing docs/config examples mention the new tags option and default behavior.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented configurable AnkiConnect tagging across SubMiner card creation and update workflows. Added `ankiConnect.tags` config (default `['SubMiner']`), validation/fallback in config parsing, AnkiConnect client support for `addNote` tags + `addTags`, and integration hooks so created/updated/merged cards receive configured tags. Updated docs and regenerated config examples; config test suite passes.
<!-- SECTION:FINAL_SUMMARY:END -->
