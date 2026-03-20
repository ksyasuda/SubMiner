---
id: TASK-192
title: Fix stale anime cover art after AniList reassignment
status: Done
assignee:
  - codex
created_date: '2026-03-20 00:12'
updated_date: '2026-03-20 00:14'
labels:
  - stats
  - immersion-tracker
  - anilist
milestone: m-1
dependencies: []
references:
  - src/core/services/immersion-tracker-service.ts
  - src/core/services/immersion-tracker/query.ts
  - src/core/services/immersion-tracker-service.test.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Fix the stats anime-detail cover image path so reassigning an anime to a different AniList entry replaces the stored cover art bytes instead of keeping the previous image blob under updated metadata.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Reassigning an anime to a different AniList entry stores the new cover art bytes for that anime's videos
- [x] #2 Shared blob deduplication still works when multiple videos in the anime use the same new cover image
- [x] #3 Focused regression coverage proves stale cover blobs are replaced on reassignment
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a failing regression test that reassigns an anime twice with different downloaded cover bytes and asserts the resolved cover updates.
2. Update cover-art upsert logic so new blob bytes generate a new shared hash instead of reusing an existing hash for the row.
3. Run the focused immersion tracker service test file and record the result.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-03-20: Created during live debugging of a user-reported stale anime profile picture after changing the AniList entry from the stats UI.
2026-03-20: Root cause was in `upsertCoverArt(...)`. When a row already had `cover_blob_hash`, a later AniList reassignment with a freshly downloaded cover reused the existing hash instead of hashing the new bytes, so the blob store kept serving the old image while metadata changed.
2026-03-20: Added a regression in `src/core/services/immersion-tracker-service.test.ts` that reassigns the same anime twice with different fetched image bytes and asserts the resolved anime cover changes to the second blob while both videos still deduplicate to one shared hash.
2026-03-20: Fixed `src/core/services/immersion-tracker/query.ts` so incoming cover blob bytes compute a fresh hash before falling back to an existing row hash. Existing hashes are now reused only when no new bytes were fetched.
2026-03-20: Verification commands run:
  - `bun test src/core/services/immersion-tracker-service.test.ts`
  - `bash .agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh src/core/services/immersion-tracker/query.ts src/core/services/immersion-tracker-service.test.ts`
  - `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane core src/core/services/immersion-tracker/query.ts src/core/services/immersion-tracker-service.test.ts`
2026-03-20: Verification results:
  - focused service test: passed
  - verifier lane selection: `core`
  - verifier result: passed (`bun run typecheck`, `bun run test:fast`)
  - verifier artifacts: `.tmp/skill-verification/subminer-verify-20260320-001433-IZLFqs/`
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed stale anime cover art after AniList reassignment by correcting cover-blob hash replacement in the immersion tracker storage layer. Reassignments now store the new fetched image bytes instead of reusing the previous blob hash from the row, while still deduplicating the updated image across videos in the same anime.

Added focused regression coverage that reproduces the exact failure mode: same anime reassigned twice with different cover downloads, with the second image expected to replace the first. Verified with the touched service test file plus the SubMiner `core` verification lane.
<!-- SECTION:FINAL_SUMMARY:END -->
