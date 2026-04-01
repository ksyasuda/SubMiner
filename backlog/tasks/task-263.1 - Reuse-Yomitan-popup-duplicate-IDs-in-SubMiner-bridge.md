---
id: TASK-263.1
title: Reuse Yomitan popup duplicate IDs in SubMiner bridge
status: Done
assignee: []
created_date: '2026-03-31 22:15'
updated_date: '2026-03-31 22:21'
labels:
  - anki
  - kiku
  - yomitan
dependencies: []
parent_task_id: TASK-263
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Thread Yomitan popup/search duplicate note IDs through the existing SubMiner bridge so Kiku/manual grouping can reuse the same duplicate context that already drives the Add duplicate button. Implement and test against the vendored Yomitan copy first; do not rely on upstreamed fork changes yet.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Vendored Yomitan bridge returns duplicate note IDs for popup/search mining when available
- [x] #2 SubMiner consumes the bridged duplicate IDs and prefers them for Kiku/manual grouping on the Yomitan mining path
- [x] #3 Regression tests cover the popup/search bridge payload and duplicate-id reuse behavior
- [x] #4 No commit is made for vendored Yomitan-only changes in this repo state
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Vendored files changed locally for validation only: vendor/subminer-yomitan/ext/js/display/display-anki.js, vendor/subminer-yomitan/ext/js/comm/api.js, vendor/subminer-yomitan/ext/js/comm/anki-connect.js, vendor/subminer-yomitan/ext/js/background/backend.js. Do not commit those vendor changes in this repo; port them to the fork instead.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Vendored Yomitan popup/search mining now precomputes duplicate note IDs, sends them to the SubMiner Anki proxy as private addNote metadata, and still returns note/duplicate data through the parser bridge. The proxy strips the private metadata before forwarding to upstream AnkiConnect, associates the duplicate IDs with the created note before auto-enrichment begins, and SubMiner also records the bridge result as a secondary cache path. Verified with bun test src/anki-integration/duplicate.test.ts src/anki-integration/card-creation.test.ts src/anki-integration/field-grouping.test.ts src/anki-integration/anki-connect-proxy.test.ts src/core/services/tokenizer/yomitan-parser-runtime.test.ts and bun run typecheck.
<!-- SECTION:FINAL_SUMMARY:END -->
