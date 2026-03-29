---
id: TASK-246
title: Migrate Discord Rich Presence to maintained RPC wrapper
status: Done
assignee: []
created_date: '2026-03-29 08:17'
updated_date: '2026-03-29 08:22'
labels:
  - dependency
  - discord
  - presence
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Replace the deprecated Discord Rich Presence wrapper with a maintained JavaScript alternative while preserving the current IPC-based presence behavior in the Electron main process.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 The app no longer depends on `discord-rpc`
- [x] #2 Discord Rich Presence still logs in and publishes activity updates from the main process
- [x] #3 Existing Discord presence tests continue to pass or are updated to cover the new client API
- [x] #4 The change is documented in the release notes or changelog fragment
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Completed:
- Swapped the app's Discord RPC dependency from `discord-rpc` to `@xhayper/discord-rpc`.
- Extracted the client adapter into `src/main/runtime/discord-rpc-client.ts` so the main process can keep using a small wrapper around the maintained library.
- Added `src/main/runtime/discord-rpc-client.test.ts` to verify the adapter forwards login/activity/clear/destroy calls through `client.user`.
- Documented the dependency swap in `CHANGELOG.md`, `release/release-notes.md`, and `docs-site/changelog.md`.

Verification:
- `bunx bun@1.3.5 test src/main/runtime/discord-rpc-client.test.ts src/core/services/discord-presence.test.ts`
- `bunx bun@1.3.5 run changelog:lint`
- `bunx bun@1.3.5 run changelog:check --version 0.10.0`
- `bunx bun@1.3.5 run docs:test`
- `bunx bun@1.3.5 run docs:build`

Notes:
- The existing release prep artifacts for v0.10.0 were kept intact and updated in place.
- No README change was needed for this dependency swap.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Replaced the deprecated `discord-rpc` dependency with the maintained `@xhayper/discord-rpc` wrapper while preserving the main-process rich presence flow. Added a focused runtime wrapper test, kept the existing Discord presence service tests green, and documented the dependency swap in the release notes and changelog.
<!-- SECTION:FINAL_SUMMARY:END -->
