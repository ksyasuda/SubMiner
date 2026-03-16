---
id: TASK-117
title: Prepare initial Windows release docs and version bump
status: Done
assignee:
  - codex
created_date: '2026-03-08 15:17'
updated_date: '2026-03-16 05:13'
labels:
  - release
  - docs
  - windows
dependencies: []
references:
  - package.json
  - README.md
  - ../subminer-docs/installation.md
  - ../subminer-docs/usage.md
  - ../subminer-docs/changelog.md
priority: medium
ordinal: 53500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Prepare the initial packaged Windows release by bumping the app version and refreshing the release-facing README/backlog/docs surfaces so install and direct-command guidance no longer reads Linux-only.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 App version is bumped for the Windows release cut
- [x] #2 README and sibling docs describe Windows packaged installation alongside Linux/macOS guidance
- [x] #3 Backlog records the release-doc/version update with the modified references
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Bump the package version for the release cut.
2. Update the root README install/start guidance to mention Windows packaged builds.
3. Patch the sibling docs repo installation, usage, and changelog pages for the Windows release.
4. Record the work in Backlog and run targeted verification on the touched surfaces.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
The public README still advertised Linux/macOS only, while the sibling docs had Windows-specific runtime notes but no actual Windows install section and several direct-command examples still assumed `SubMiner.AppImage`.

Bumped `package.json` to `0.5.0`, expanded the README platform/install copy to include Windows, added a Windows install section to `../subminer-docs/installation.md`, clarified in `../subminer-docs/usage.md` that direct packaged-app examples use `SubMiner.exe` on Windows, and added a `v0.5.0` changelog entry covering the initial Windows release plus the latest overlay behavior polish.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Prepared the initial Windows release documentation pass and version bump. `package.json` now reports `0.5.0`. The root `README.md` now advertises Linux, macOS, and Windows support, includes Windows packaged-install guidance, and clarifies first-launch behavior across platforms. In the sibling docs repo, `installation.md` now includes a dedicated Windows install section, `usage.md` explains that direct packaged-app examples use `SubMiner.exe` on Windows, and `changelog.md` now includes the `v0.5.0` release notes for the initial Windows build and recent overlay behavior changes.

Verification: targeted `bun run tsc --noEmit -p tsconfig.typecheck.json` in the app repo and `bun run docs:build` in `../subminer-docs`.
<!-- SECTION:FINAL_SUMMARY:END -->
