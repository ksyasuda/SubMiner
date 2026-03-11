---
id: TASK-155
title: Move user docs site back into main repo
status: Done
assignee: []
created_date: '2026-03-10 19:20'
updated_date: '2026-03-10 19:38'
labels: []
dependencies: []
priority: medium
ordinal: 15500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Move the standalone VitePress docs site from the sibling `../subminer-docs` checkout back into the main `SubMiner` repo so docs can be updated alongside code and local tooling can reference one repository.

Scope:

- import the tracked docs-site source into a dedicated in-repo subdirectory
- update scripts/tests/docs instructions that assume a sibling `../subminer-docs` checkout
- preserve Cloudflare Pages deployability from a repo subdirectory
- verify the app repo and docs site both still build/test from the new layout

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 The user-facing VitePress docs source lives inside the `SubMiner` repo in a dedicated subdirectory.
- [x] #2 First-party scripts/tests/docs no longer require `../subminer-docs` for normal operation.
- [x] #3 In-repo docs instructions include the Cloudflare Pages subdirectory deploy settings.
- [x] #4 Verification covers the relocated docs site build/tests plus affected app-repo checks.

<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Imported the VitePress site into `docs-site/` inside the main repo and updated project instructions, docs contributor guidance, generator logic, and regression tests to treat that in-repo directory as the docs source of truth.

Added root proxy scripts for `docs:dev`, `docs:build`, `docs:preview`, and `docs:test`, repointed config-example generation to `docs-site/public/config.example.jsonc`, switched docs edit links to the main `SubMiner` repo, and documented the Cloudflare Pages subdirectory settings (`docs-site` root, `.vitepress/dist` output, `docs-site/**` watch path).

Verified with `bun run format:check:src`, `bun run typecheck`, `bun run docs:test`, `bun run docs:build`, `bun run test:config:src`, and `bun run test:fast`.

<!-- SECTION:FINAL_SUMMARY:END -->
