---
id: TASK-251
title: 'Docs: add subtitle sidebar and Jimaku integration pages'
status: Done
assignee: []
created_date: '2026-03-29 22:36'
updated_date: '2026-03-31 19:37'
labels:
  - docs
dependencies: []
priority: medium
ordinal: 164500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Track the docs-site update that adds a dedicated subtitle sidebar page, links Jimaku integration from the homepage/config docs, and refreshes the docs-site theme styling used by those pages.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 docs-site nav includes a Subtitle Sidebar entry
- [x] #2 Subtitle Sidebar page documents layout, shortcut, and config options
- [x] #3 Jimaku integration page and configuration docs link to the new docs page
- [x] #4 Changelog fragment exists for the user-visible docs release note
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added the subtitle sidebar docs page and nav entry, linked Jimaku integration from the homepage/config docs, refreshed docs-site styling tokens, and recorded the release note fragment. Verified with `bun run changelog:lint`, `bun run docs:test`, `bun run docs:build`, and `bun run build`. Full repo test gate still has pre-existing failures in `bun run test:fast` and `bun run test:env` unrelated to these docs changes.
<!-- SECTION:FINAL_SUMMARY:END -->
