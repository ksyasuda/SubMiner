---
id: TASK-63
title: Add runtime toggle to log selected Yomitan token groups
status: Done
assignee: []
created_date: '2026-02-16 23:38'
updated_date: '2026-02-18 04:11'
labels: []
dependencies: []
priority: low
ordinal: 22000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Provide an in-app debug toggle that logs the selected Yomitan token grouping for each subtitle parse so users can verify token boundaries live without rebuilding.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 A runtime option exists to enable/disable Yomitan group debug logging without app restart.
- [x] #2 When enabled, subtitle tokenization logs the selected Yomitan grouped tokens (with enough detail to verify boundaries/headwords).
- [x] #3 When disabled, no additional Yomitan group debug logs are emitted.
- [x] #4 Related tests/build pass for touched modules.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a boolean runtime option for Yomitan-group debug logging in the centralized runtime option registry and expose it in config metadata.
2. Extend tokenizer dependency wiring so main runtime can pass the current toggle value to tokenization without restart.
3. Log selected Yomitan token groups (surface/headword/reading/span) only when the toggle is enabled.
4. Add tests for registry presence and enabled/disabled logging behavior, then run build/tests.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added runtime option `anki.debugYomitanGroups` (`Debug Yomitan Groups`) with default `false`, mapped to `ankiConnect.behavior.debugYomitanGroups`.

Wired `main.ts` tokenizer deps to read the runtime option value live via `RuntimeOptionsManager`, with config fallback.

Implemented conditional tokenizer logging (`Selected Yomitan token groups`) in `tokenizer-service` and covered enabled/disabled behavior with unit tests.

Validation run: `pnpm run build && node dist/core/services/tokenizer-service.test.js && node --test dist/config/config.test.js` (all passing).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented a live runtime debug toggle to inspect Yomitan token grouping. Added `anki.debugYomitanGroups` to the runtime option registry and config defaults, wired it through `main.ts` into tokenizer deps, and added conditional logging in tokenizer parsing that emits selected groups with surface/headword/reading/span for each parsed subtitle. Logging is gated by the toggle and disabled by default. Added tests for runtime registry presence and tokenizer logging on/off behavior, then validated with build + tokenizer + config tests.
<!-- SECTION:FINAL_SUMMARY:END -->
