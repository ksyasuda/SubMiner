---
id: TASK-112
title: Fix docs demo video thumbnail cache staleness
status: Done
assignee:
  - codex
created_date: '2026-02-23 03:41'
updated_date: '2026-02-23 03:43'
labels:
  - bug
  - docs
  - media
dependencies: []
priority: medium
ordinal: 112000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Docs home page demo continues to show an old thumbnail after replacing the demo video file in place. The page currently reuses static media URLs (`/assets/minecard.webm`, `/assets/minecard.mp4`, `/assets/demo-poster.jpg`), so browser/CDN cache can keep stale poster/video preview data.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Add a failing regression check that asserts cache-busted URL usage for docs home demo media references.
2. Update docs home demo media URLs to include a shared asset version token for poster/source/fallback paths.
3. Run the regression check and docs build validation.
4. Record fix notes and move task status to Done.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Docs home video references include a deterministic cache-busting token so direct asset replacement does not serve stale preview data.
- [x] #2 Poster/source/fallback links in the demo block use the same token to avoid mixed old/new media.
- [x] #3 Validation commands for the new regression check and docs build pass.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
- Generated a fresh thumbnail from `docs/public/assets/minecard.webm` into `docs/public/assets/minecard-poster.jpg`.
- Updated `docs/index.md` demo block to use a shared `demoAssetVersion` token across poster, webm/mp4 sources, and gif fallback links.
- Follow-up: regenerated `minecard-poster.jpg` at exact `00:00:12` from the webm and bumped token to `demoAssetVersion = '20260223-2'`.
- Added regression check `docs/index.assets.test.ts` to ensure demo media references keep shared cache-busting URL patterns.
- Validation:
  - `bun test docs/index.assets.test.ts`
  - `bun run docs:build`
<!-- SECTION:NOTES:END -->
