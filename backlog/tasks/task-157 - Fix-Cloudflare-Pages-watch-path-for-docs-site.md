---
id: TASK-157
title: Fix Cloudflare Pages watch path for docs-site
status: Done
assignee: []
created_date: '2026-03-10 20:15'
updated_date: '2026-03-16 05:13'
labels:
  - docs-site
  - cloudflare
dependencies: []
priority: medium
ordinal: 98500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Cloudflare Pages skipped a docs-site deployment after the docs repo moved into the main `SubMiner` repository. The documented/configured watch path uses `docs-site/**`, but Pages monorepo watch paths use a single `*` wildcard pattern. Correct the documented setting and leave a regression test so future repo moves or docs rewrites do not reintroduce the bad pattern.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Docs contributor guidance points Cloudflare Pages watch paths at `docs-site/*`, not `docs-site/**`.
- [x] #2 Regression coverage fails if the docs revert to the incorrect watch-path string.
- [x] #3 Implementation notes record that the Cloudflare dashboard setting must be updated manually and the docs deploy retriggered.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added docs regression coverage in `docs-site/docs-sync.test.ts` for the Pages watch-path string, then corrected the Cloudflare Pages instructions in `docs-site/README.md` and `docs-site/development.md`.

Manual follow-up still required outside git: update the Cloudflare Pages project include path from `docs-site/**` to `docs-site/*`, then trigger a fresh deployment against `main`.
<!-- SECTION:NOTES:END -->
