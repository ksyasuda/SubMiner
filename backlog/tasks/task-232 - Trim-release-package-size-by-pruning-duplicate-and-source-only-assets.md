---
id: TASK-232
title: Trim release package size by pruning duplicate and source-only assets
status: Done
assignee:
  - '@codex'
created_date: '2026-03-24 12:05'
updated_date: '2026-03-24 12:30'
labels:
  - release
  - packaging
priority: medium
ordinal: 54700
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/package.json
  - /home/sudacode/projects/japanese/SubMiner/src/release-workflow.test.ts
  - /home/sudacode/projects/japanese/SubMiner/src/core/services/texthooker.ts
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Reduce packaged release artifact size without changing user-visible functionality by pruning files that are duplicated between `app.asar` and `extraResources`, excluding source/test/doc-only trees from Electron packaging, and trimming obviously non-runtime vendored payload.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Electron packaging excludes repo content that is source-only, test-only, docs-only, or duplicated by `extraResources`.
- [x] #2 Release packaging tests cover the new exclusion rules.
- [x] #3 Verification includes at least targeted release-packaging tests and one packaging-oriented validation step.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Completed scope:
- Exclude `assets`, `plugin`, and `vendor/yomitan-jlpt-vocab` from `files` because they are already staged via `extraResources`.
- Exclude `dist` sourcemaps/tests, repo docs/tests/packaging metadata, and stats source leftovers from `files`.
- Exclude non-runtime `vendor/texthooker-ui` payload such as `public/`, `.vscode/`, and package metadata.
- Exclude Linux musl libsql binary from packaged app payload for AppImage-focused savings.

Verification:
- `bun test src/release-workflow.test.ts`
- `bun run build`
- `node_modules/.bin/electron-builder --linux dir --publish never`
- `node_modules/.bin/electron-builder --linux AppImage --publish never`

Observed result:
- `release/linux-unpacked/resources/app.asar` dropped from about `100 MB` to `29 MB`.
- `release/SubMiner-0.9.0.AppImage` dropped from about `256 MB` to `194 MB`.

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Trimmed Electron packaging so release artifacts no longer bundle duplicated `extraResources`, source/test/doc-only repo content, non-runtime `texthooker-ui` files, or the Linux musl libsql binary. Added release-packaging regression coverage and verified the Linux package shrink with fresh local builds.

<!-- SECTION:FINAL_SUMMARY:END -->
