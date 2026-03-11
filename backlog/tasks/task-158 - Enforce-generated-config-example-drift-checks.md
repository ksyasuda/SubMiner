---
id: TASK-158
title: Enforce generated config example drift checks
status: Done
assignee: []
created_date: '2026-03-10 20:35'
updated_date: '2026-03-10 20:35'
labels:
  - config
  - docs-site
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
The generated `config.example.jsonc` artifact is covered by generation-path tests, but there is no hard gate that fails when the checked-in example drifts from the canonical template. The in-repo docs-site copy can also drift silently.

Scope:

- add a first-party verification path that compares generated config-example content against committed artifacts
- fail fast when repo-root `config.example.jsonc` is stale or missing
- fail fast when `docs-site/public/config.example.jsonc` is stale or missing, when the docs site exists
- wire the verification into the normal gate and release flow
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->
- [x] #1 Automated verification fails when repo-root `config.example.jsonc` is missing or stale.
- [x] #2 Automated verification fails when in-repo docs-site `public/config.example.jsonc` is missing or stale, when docs-site exists.
- [x] #3 CI/release or equivalent project gates run the verification automatically.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added `src/verify-config-example.ts`, which renders the canonical config template and compares it against the checked-in repo-root `config.example.jsonc` plus `docs-site/public/config.example.jsonc` when the docs site exists.

Wired the new verification into `package.json` as `bun run verify:config-example`, added regression coverage for missing and stale artifacts, and enforced the new check in both CI and release workflows.

Regenerated the checked-in config example artifacts so the new gate passes in the repo-local docs-site layout, and documented the release-step expectation in `docs/RELEASING.md`.
<!-- SECTION:FINAL_SUMMARY:END -->
