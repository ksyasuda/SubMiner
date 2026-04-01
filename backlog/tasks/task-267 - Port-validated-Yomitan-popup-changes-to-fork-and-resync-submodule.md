---
id: TASK-267
title: Port validated Yomitan popup changes to fork and resync submodule
status: Done
assignee: []
created_date: '2026-04-01 03:30'
updated_date: '2026-04-01 03:33'
labels:
  - yomitan
  - submodule
  - git
  - integration
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Take the locally validated Yomitan popup/bridge changes from the vendored copy, apply them to the standalone `../subminer-yomitan` fork, verify the fork, push the fork commit, then reset the vendored working tree in SubMiner and update the submodule pointer to the pushed fork commit.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Standalone `../subminer-yomitan` contains the validated popup/bridge changes and passes the relevant regression test
- [x] #2 The fork commit is pushed to its configured remote branch
- [x] #3 SubMiner vendored Yomitan working tree is reset and the submodule pointer is updated to the pushed fork commit
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Applied the validated popup/bridge changes from the vendored Yomitan copy into `../subminer-yomitan`, added the focused async-save regression test there, installed fork deps, and verified with `npx vitest run test/display-anki-save.test.js`. Committed the fork changes as `feat: preserve async popup save state and duplicate metadata`, rebased onto the updated remote `main`, and pushed commit `69620abc` to `origin/main`. Then reset the vendored submodule working tree in SubMiner, checked it out at `69620abc`, and left the superproject with the submodule pointer updated from `3c9ee577` to `69620abc`.
<!-- SECTION:FINAL_SUMMARY:END -->
