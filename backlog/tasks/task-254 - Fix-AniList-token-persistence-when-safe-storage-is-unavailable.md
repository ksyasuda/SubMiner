---
id: TASK-254
title: Fix AniList token persistence when safe storage is unavailable
status: Done
assignee:
  - codex
created_date: '2026-03-30 02:10'
updated_date: '2026-03-30 02:20'
labels:
  - bug
  - anilist
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/anilist/anilist-token-store.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/main/runtime/anilist-setup.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/main/runtime/anilist-token-refresh.ts
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
AniList login currently appears to succeed during setup, but some environments cannot persist the token because Electron safeStorage is unavailable or unusable. On the next app start, AniList tracking cannot load the token and re-prompts the user to set up AniList again. Align AniList token persistence with the intended login UX so a token the user already saved is reused on later sessions.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Saved encrypted AniList token is reused on app-ready startup without reopening setup.
- [x] #2 AniList startup no longer attempts to open the setup BrowserWindow before Electron is ready.
- [x] #3 AniList auth/runtime tests cover stored-token reuse and the missing-token startup path that previously triggered pre-ready setup attempts.
- [x] #4 AniList token storage remains encrypted-only; no plaintext fallback is introduced.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add regression tests for AniList startup auth refresh so a stored encrypted token is reused without opening setup, and for the missing-token path so setup opening is deferred safely until the app can actually show a window.
2. Update AniList startup/auth runtime to separate token resolution from setup-window prompting, and gate prompting on app readiness instead of attempting BrowserWindow creation during early startup.
3. Preserve encrypted-only storage semantics in anilist-token-store; do not add plaintext fallback. If stored-token load fails, keep logging/diagnostics intact.
4. Run targeted AniList runtime/token tests, then summarize root cause and verification results.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Investigated AniList auth persistence flow. Current setup path treats callback token as saved even when anilist-token-store refuses persistence because safeStorage is unavailable. Jellyfin token store already uses plaintext fallback in this environment class, which is a likely model for the AniList fix.

Confirmed from local logs that safeStorage was explicitly unavailable on 2026-03-23 due macOS Keychain lookup failure with NSOSStatusErrorDomain Code=-128 userCanceledErr. Current environment also has an encrypted AniList token file at /Users/sudacode/.config/SubMiner/anilist-token-store.json updated 2026-03-29 18:49, so safeStorage did work recently for save. Repeated AniList setup prompts on 2026-03-29/30 correlate more strongly with startup auth flow deciding no token is available and opening setup immediately; logs show repeated 'Loaded AniList manual token entry page' and several 'Failed to refresh AniList client secret state during startup' errors with 'Cannot create BrowserWindow before app is ready'. No recent log lines indicate safeStorage.isEncryptionAvailable() false after 2026-03-23.

Implemented encrypted-only startup fix by adding an allowSetupPrompt control to AniList token refresh and disabling setup-window prompting for the early pre-ready startup refresh in main.ts. App-ready reloadConfig still performs the normal prompt-capable refresh after Electron is ready. Added regression tests for stored-token reuse and prompt suppression when startup explicitly disables prompting.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Root cause was a redundant early AniList auth refresh during startup. The app refreshed AniList auth once before Electron was ready and again during app-ready config reload. When the early refresh could not resolve a token, it tried to open the AniList setup window immediately, which produced the observed 'Cannot create BrowserWindow before app is ready' failures and repeated setup prompts. The fix keeps token storage encrypted-only, teaches AniList auth refresh to optionally suppress setup-window prompting, and uses that suppression for the early startup refresh. App-ready startup still performs the normal prompt-capable refresh once Electron is ready, so saved encrypted tokens are reused without reopening setup and missing-token setup only happens at a safe point. Verified with targeted AniList auth tests, typecheck, test:fast, test:env, build, and test:smoke:dist.
<!-- SECTION:FINAL_SUMMARY:END -->
