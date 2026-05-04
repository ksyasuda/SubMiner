---
id: TASK-335
title: Fix Linux AniList setup gate using stored keyring token
status: Done
assignee: []
created_date: '2026-05-04 05:26'
updated_date: '2026-05-04 05:30'
labels:
  - anilist
  - bug
  - linux
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
AniList setup page reopens on Linux video launch even when the token exists in secret storage and post-watch updates can use it. Investigate setup gating versus update token refresh paths and make them agree on stored-token availability.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launching a video on Linux with an AniList token available in secret storage does not show the AniList setup page just because config accessToken is empty.
- [x] #2 If secret storage load fails, setup/errors surface the underlying storage problem instead of behaving like an empty token.
- [x] #3 Regression coverage exercises the setup-gate token availability path and preserves post-watch update token behavior.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Patched AniList setup callback to require successful token persistence before caching/closing the setup flow. Patched config reload auth refresh to pass allowSetupPrompt:false so normal startup/playback reloads do not open AniList setup UI. Added regression coverage around persistence failure and non-prompting config refresh.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed AniList setup/login flow so failed encrypted token persistence no longer reports success or seeds only an in-memory token. Config reload now refreshes AniList auth state without opening the setup window during playback, reducing repeated Linux setup prompts when safeStorage/keyring resolution fails.
<!-- SECTION:FINAL_SUMMARY:END -->
