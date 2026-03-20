---
id: TASK-156
title: Fix docs-site Plausible geo attribution through analytics worker
status: Done
assignee: []
created_date: '2026-03-11 02:19'
updated_date: '2026-03-16 05:13'
labels:
  - docs-site
  - analytics
dependencies: []
priority: medium
ordinal: 99500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Investigate and fix missing city/region data for docs.subminer.moe visits in Plausible. The docs site already proxies analytics events to worker.subminer.moe, so the remaining work is to verify and correct the worker-side forwarding contract Plausible needs for geolocation.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 The analytics worker forwards Plausible event requests in a way that preserves the original client IP information needed for location attribution.
- [ ] #2 The docs-site analytics flow remains proxied through worker.subminer.moe after the fix.
- [ ] #3 Coverage or documentation records the worker-side header/forwarding requirement for Plausible geo reporting.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Worker implementation lives in `/home/sudacode/projects/blog-proxy`, not in the SubMiner repo.

Patched `worker.js` to forward client IP via `x-forwarded-for`/`x-real-ip` from `cf-connecting-ip` (fallbacks retained) and added `worker.test.js` regression coverage.

Local verification in `blog-proxy`: `node --test worker.test.js` passes. Deployment not performed in this session.
<!-- SECTION:NOTES:END -->
