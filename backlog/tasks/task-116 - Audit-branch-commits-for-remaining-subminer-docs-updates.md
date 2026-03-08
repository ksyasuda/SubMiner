---
id: TASK-116
title: Audit branch commits for remaining subminer-docs updates
status: Done
assignee:
  - codex
created_date: '2026-03-08 00:46'
updated_date: '2026-03-08 00:48'
labels:
  - docs
dependencies: []
references:
  - ../subminer-docs/installation.md
  - ../subminer-docs/troubleshooting.md
  - src/core/services/yomitan-extension-paths.ts
  - scripts/build-yomitan.mjs
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Review recent `yomitan-fork` commits against the sibling `subminer-docs` repo, identify any concrete documentation drift that remains after the earlier contributor-doc updates, and patch the docs for behavior/tooling changes that are now outdated or misleading.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Reviewed recent branch commits for user-facing or contributor-facing changes that may require docs updates
- [x] #2 Updated `subminer-docs` pages where branch changes introduced concrete doc drift
- [x] #3 Verified the docs site still builds after the updates
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Review branch commit themes against `subminer-docs` and identify only concrete drift introduced by recent workflow/runtime changes.
2. Patch docs for the Yomitan submodule build workflow, updated source-build prerequisites, and current runtime Yomitan search paths/manual fallback path.
3. Rebuild the docs site to verify the updated pages render cleanly.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Concrete remaining drift after commit audit: installation/development docs still understate the Node/npm + submodule requirements for the Yomitan build flow, and troubleshooting still points at obsolete `vendor/yomitan` / `extensions/yomitan` paths.

Audited branch commits against subminer-docs coverage. Existing docs already cover first-run setup, texthooker startup/annotated websocket config, AniList merged character dictionaries, configurable collapsible sections, and subtitle name highlighting. Patched remaining drift around source-build prerequisites and Yomitan build/install paths in installation.md, development.md, and troubleshooting.md. Verified with `bun run docs:build` in ../subminer-docs.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Audited branch commits for missing documentation updates in ../subminer-docs. Updated installation, development, and troubleshooting docs to match the current Yomitan submodule build flow, source-build prerequisites, and runtime extension search/manual fallback paths. Confirmed other recent branch features were already documented and rebuilt the docs site successfully.
<!-- SECTION:FINAL_SUMMARY:END -->
