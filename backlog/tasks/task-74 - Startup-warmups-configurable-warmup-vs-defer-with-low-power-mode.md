---
id: TASK-74
title: 'Startup warmups: configurable warmup vs defer with low-power mode'
status: In Progress
assignee: []
created_date: '2026-02-27 21:05'
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
  - src/main.ts
  - src/config/config.test.ts
  - src/main/runtime/startup-warmups.test.ts
priority: medium
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
<!-- SECTION:FINAL_SUMMARY:END -->
