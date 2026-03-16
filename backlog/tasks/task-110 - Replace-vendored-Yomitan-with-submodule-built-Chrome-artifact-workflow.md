---
id: TASK-110
title: Replace vendored Yomitan with submodule-built Chrome artifact workflow
status: Done
assignee: []
created_date: '2026-03-07 11:05'
updated_date: '2026-03-16 05:13'
labels:
  - yomitan
  - build
  - release
dependencies: []
priority: high
ordinal: 10010
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Replace the checked-in `vendor/yomitan` release tree with a `subminer-yomitan` git submodule. Build Yomitan from source, extract the Chromium zip artifact into a stable local build directory, and make SubMiner dev/runtime/tests/release packaging load that extracted extension instead of the source tree or vendored files.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Repo tracks Yomitan as a git submodule instead of committed extension files under `vendor/yomitan`.
- [x] #2 SubMiner has a reproducible build/extract step that produces a local Chromium extension directory from `subminer-yomitan`.
- [x] #3 Dev/runtime/tests resolve the extracted build output as the default Yomitan extension path.
- [x] #4 Release packaging includes the extracted Chromium extension files instead of the old vendored tree.
- [x] #5 Docs and verification commands reflect the new workflow.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Replaced the checked-in `vendor/yomitan` extension tree with a `vendor/subminer-yomitan` git submodule and added a reproducible `bun run build:yomitan` workflow that builds `yomitan-chrome.zip`, extracts it into `build/yomitan`, and reuses a source-state stamp to skip redundant rebuilds. Runtime path resolution, helper CLIs, Yomitan integration tests, packaging, CI cache keys, and README source-build notes now all target that generated artifact instead of the old vendored files.

Verification:

- `bun run build:yomitan`
- `bun test src/core/services/yomitan-extension-paths.test.ts src/core/services/yomitan-structured-content-generator.test.ts src/yomitan-translator-sort.test.ts`
- `bun run typecheck`
- `bun run build`
- `bun run test:core:src`
<!-- SECTION:FINAL_SUMMARY:END -->
