---
id: TASK-168
title: Document immersion stats dashboard and config
status: Done
assignee:
  - codex
created_date: '2026-03-12 22:53'
updated_date: '2026-03-12 22:53'
labels:
  - docs
  - immersion
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Refresh user-facing docs for the new immersion stats dashboard so README, docs-site pages, changelog notes, and generated config examples describe how to access the dashboard and which `stats.*` settings control it.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 README mentions the new stats surface in product-facing feature/docs copy.
- [x] #2 Docs explain how to access the stats dashboard in-app and via localhost, and document the `stats` config block.
- [x] #3 Changelog/release-note input includes the new stats dashboard.
- [x] #4 Generated config examples include the new `stats` section.

<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Updated README and the docs-site immersion/config/mining/shortcut/homepage copy to describe the new stats dashboard, including the overlay toggle (`stats.toggleKey`, default `Backquote`) and the localhost browser UI (`http://127.0.0.1:5175` by default).

Added a changelog fragment for the stats dashboard release notes and extended the config template sections so regenerated `config.example.jsonc` artifacts now include the `stats` block.

Verified with `bun run test:config`, `bun run generate:config-example`, `bun run docs:test`, `bun run docs:build`, and `bun run changelog:lint`.

<!-- SECTION:FINAL_SUMMARY:END -->
