---
id: TASK-74
title: 'Startup warmups: configurable warmup vs defer with low-power mode'
status: Done
assignee: []
created_date: '2026-02-27 21:05'
updated_date: '2026-03-01 04:14'
labels: []
dependencies: []
references:
  - src/types.ts
  - src/config/definitions/defaults-core.ts
  - src/config/definitions/options-core.ts
  - src/config/definitions/template-sections.ts
  - src/config/resolve/core-domains.ts
  - src/main/runtime/startup-warmups.ts
  - src/main/runtime/startup-warmups-main-deps.ts
  - src/main/runtime/composers/mpv-runtime-composer.ts
  - src/core/services/startup.ts
  - src/main.ts
  - src/config/config.test.ts
  - src/main/runtime/startup-warmups.test.ts
  - src/main/runtime/startup-warmups-main-deps.test.ts
  - src/core/services/app-ready.test.ts
priority: medium
ordinal: 7000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Add startup warmup controls to allow per-integration warmup or deferred first-use loading.

Scope:

- New config section `startupWarmups` with toggles for `mecab`, `yomitanExtension`, `subtitleDictionaries`, and `jellyfinRemoteSession`.
- New `startupWarmups.lowPowerMode` policy: defer everything except Yomitan extension.
- Keep default behavior as full warmup.
- Ensure deferred integrations lazy-load on first real usage path.
- Add test coverage for config parsing/defaults and warmup scheduling behavior.
<!-- SECTION:DESCRIPTION:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Implemented:

- Added `startupWarmups` to config types/defaults/options/template/resolve.
- Warmup scheduler now uses per-integration gating functions.
- Low-power mode now defers MeCab, subtitle dictionaries, and Jellyfin remote session warmups while still warming Yomitan extension.
- Tokenization path guarantees lazy first-use init for deferred dependencies (Yomitan extension, MeCab when missing, subtitle dictionaries).
- Added/updated tests across config and runtime warmup modules.

Validation:

- `bun run test:config:src`
- `bun run test:core:src`
- `tsc --noEmit`

Follow-up updates:

- Startup now triggers warmups earlier in app-ready flow (right after config validation/log-level setup) instead of waiting for initial args/overlay actions. Goal: tokenization warmup is already done or mostly done by first visible-subs toggle.
- Tokenization warmup scheduling consolidated as `subtitle-tokenization` stage; when enabled by toggles, it runs Yomitan extension first, then MeCab/dictionary warmups.
- Added per-stage debug logs for warmup progress and skip reasons:
  - `stage start/ready: yomitan-extension`
  - `stage start/ready: mecab`
  - `stage start/ready: subtitle-dictionaries`
  - `stage start/ready: jellyfin-remote-session`
  - `stage skipped: jellyfin-remote-session (disabled|auto-connect off)`
- Added regression tests for stage-level logging and earlier startup ordering:
  - `src/main/runtime/startup-warmups.test.ts`
  - `src/main/runtime/startup-warmups-main-deps.test.ts`
  - `src/core/services/app-ready.test.ts`
  <!-- SECTION:FINAL_SUMMARY:END -->
