---
id: TASK-21
title: Persist and restore MPV secondary subtitle visibility across app lifecycle
status: Done
assignee: []
created_date: '2026-02-13 07:59'
updated_date: '2026-02-18 04:11'
labels: []
dependencies: []
priority: high
ordinal: 46000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

When SubMiner connects to MPV, capture the current MPV `secondary-sub-visibility` value and force it off. Keep it off during SubMiner runtime regardless of overlay visibility toggles. On app shutdown (and MPV shutdown event when possible), restore MPV `secondary-sub-visibility` to the captured pre-SubMiner value.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Capture MPV `secondary-sub-visibility` once per MPV connection before overriding it.
- [x] #2 Set MPV `secondary-sub-visibility` to `no` after capture regardless of `bind_visible_overlay_to_mpv_sub_visibility`.
- [x] #3 Do not mutate/restore secondary MPV visibility as a side effect of visible overlay toggles.
- [x] #4 Restore captured secondary MPV visibility on app shutdown while MPV is connected.
- [x] #5 Attempt restore on MPV shutdown event before disconnect and clear captured state afterward.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Implemented MPV secondary subtitle visibility lifecycle management:

- Moved secondary-sub visibility capture/disable to MPV connection initialization (`getInitialState` requests `secondary-sub-visibility`, then request handler stores prior value and forces `secondary-sub-visibility=no`).
- Removed secondary-sub visibility side effects from visible overlay visibility service so overlay toggles no longer capture/restore secondary MPV state.
- Added `restorePreviousSecondarySubVisibility()` to `MpvIpcClient`, invoked on MPV `shutdown` event and from app `onWillQuitCleanup` (best effort while connected).
- Wired new dependency getter/setter in main runtime bootstrap for tracked previous secondary visibility state.
- Added unit coverage in `mpv-service.test.ts` for capture/disable and restore/clear behavior.
- Verified with `pnpm run build` and `node --test dist/core/services/mpv-service.test.js`.
<!-- SECTION:FINAL_SUMMARY:END -->
