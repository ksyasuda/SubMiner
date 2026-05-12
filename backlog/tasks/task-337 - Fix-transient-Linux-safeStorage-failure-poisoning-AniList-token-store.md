---
id: TASK-337
title: Fix transient Linux safeStorage failure poisoning AniList token store
status: Done
assignee: []
created_date: '2026-05-04 05:51'
updated_date: '2026-05-04 05:52'
labels:
  - anilist
  - bug
  - linux
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
AniList token store memoizes a false safeStorage availability result. On Linux this can happen before Electron/keyring readiness, causing later post-watch updates and setup saves to report missing login/encryption unavailable even after the keyring is available.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 A transient safeStorage unavailable result does not prevent a later stored AniList token load once encryption is available.
- [x] #2 A transient safeStorage unavailable result does not prevent a later AniList token save once encryption is available.
- [x] #3 Regression coverage protects the retry behavior.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Changed AniList token store safeStorage probe to memoize successful probes only. Failed probes now return false without poisoning later load/save attempts, covering Linux startup windows where Electron safeStorage/keyring can be unavailable before app readiness but usable later. Added regression test for transient unavailable -> available load/save retry.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed a Linux AniList auth failure where an early safeStorage/keyring miss was cached for the whole process. Stored tokens now load and setup tokens can save after GNOME libsecret becomes available later in startup.
<!-- SECTION:FINAL_SUMMARY:END -->
