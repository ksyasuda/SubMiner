---
id: TASK-102
title: Quiet default AppImage startup and implicit background launch
status: Done
assignee:
  - codex
created_date: '2026-03-06 21:20'
updated_date: '2026-03-06 21:33'
labels: []
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/main-entry-runtime.ts
  - /home/sudacode/projects/japanese/SubMiner/src/core/services/cli-command.ts
  - /home/sudacode/projects/japanese/SubMiner/src/main.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Make the packaged Linux no-arg launch path behave like a quiet background start instead of surfacing startup-only noise.

Scope:

- Treat default background entry launches as implicit `--start --background`.
- Keep the `--password-store` diagnostic out of normal startup output.
- Suppress known startup-only `node:sqlite` and `lsfg-vk` warnings for the entry/background launch path.
- Avoid noisy protocol-registration warnings during normal startup when registration is unsupported.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Initial background launch reaches the start path without logging `No running instance. Use --start to launch the app.`
- [x] #2 Default startup no longer emits the `Applied --password-store gnome-libsecret` line at normal log levels.
- [x] #3 Entry/background launch sanitization suppresses the observed `ExperimentalWarning: SQLite...` and `lsfg-vk ... unsupported configuration version` startup noise.
- [x] #4 Regression coverage documents the new startup behavior.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Normalized no-arg/password-store-only entry launches to append implicit `--start --background`, and upgraded `--background`-only entry launches to include `--start`.

Applied shared entry env sanitization before loading the main process so default startup strips the `lsfg-vk` Vulkan layer and sets `NODE_NO_WARNINGS=1`; background children keep the same sanitized env.

Downgraded startup-only protocol-registration failure logging to debug, and routed the Linux password-store diagnostic through the scoped debug logger instead of raw console output.

Verification:

- `bun test src/main-entry-runtime.test.ts src/main/runtime/anilist-setup-protocol.test.ts src/main/runtime/anilist-setup-protocol-main-deps.test.ts`
- `bun run test:fast`

Note: the final `node --experimental-sqlite --test dist/main/runtime/registry.test.js` step in `bun run test:fast` still prints Node's own experimental SQLite warning because that test command explicitly enables the feature flag outside the app entrypoint.

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Default packaged startup is now quiet and behaves like an implicit `--start --background` launch.

- No-arg AppImage entry launches now append `--start --background`, and `--background`-only launches append the missing `--start`.
- Entry/background startup sanitization now suppresses the observed `lsfg-vk` and `node:sqlite` warnings on the app launch path.
- Linux password-store and unsupported protocol-registration diagnostics now stay at debug level instead of normal startup output.
<!-- SECTION:FINAL_SUMMARY:END -->
