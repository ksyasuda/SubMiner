---
id: TASK-359
title: Update default startup and subtitle style options
status: Done
assignee:
  - Codex
created_date: '2026-05-13 06:16'
updated_date: '2026-05-13 06:28'
labels:
  - config
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Change fresh-install defaults so texthooker and associated subtitle websocket servers are disabled unless explicitly enabled, Yomitan popup auto-pause remains enabled by default, and primary/secondary subtitle style defaults match the requested font, transparent background, and stronger shadow. Keep generated example config and docs aligned with resolved defaults.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Fresh default config has texthooker startup disabled by default.
- [x] #2 Fresh default config disables associated subtitle websocket servers by default, or exposes/uses an existing config option that does so.
- [x] #3 Fresh default config keeps subtitleStyle.autoPauseVideoOnYomitanPopup enabled by default.
- [x] #4 Fresh default config uses fontFamily `Hiragino Sans, M PLUS 1, Source Han Sans JP, Noto Sans CJK JP` for primary subtitles.
- [x] #5 Fresh default config keeps secondary subtitle fontFamily as `Inter, Noto Sans, Helvetica Neue, sans-serif`.
- [x] #6 Fresh default config uses textShadow `0 2px 6px rgba(0,0,0,0.9), 0 0 12px rgba(0,0,0,0.55)` for primary and secondary subtitles.
- [x] #7 Fresh default config uses transparent subtitle backgrounds for primary and secondary subtitles.
- [x] #8 Tests and generated/docs config examples are updated for the new defaults.
- [x] #9 Fresh default config uses `#8bd5ca` for the fourth frequency band and JLPT N4 subtitle color.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Update focused config tests first for the requested fresh-install defaults.
2. Change default config values in `src/config/definitions/defaults-core.ts` and `src/config/definitions/defaults-subtitle.ts`.
3. Regenerate generated config examples with the repo generator.
4. Update docs-site configuration prose/snippets for the new defaults.
5. Run focused config/example tests and mark acceptance criteria when verified.

6. Include the requested JLPT N4 default color adjustment and verify the existing fourth frequency band default remains `#8bd5ca`.

7. Restore the secondary subtitle font default to its previous `Inter, Noto Sans, Helvetica Neue, sans-serif` stack while keeping the new primary subtitle font default.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Context: existing config already has websocket.enabled (`true | false | "auto"`) and annotationWebsocket.enabled (`boolean`), so disabling associated websockets can be handled by default changes instead of adding a new config surface.

Scope update from user: also set JLPT N4 default color to `#8bd5ca`; fourth frequency band is already `#8bd5ca` in current defaults and has existing regression coverage.

Verification: `bun run test:config`, `bun run docs:test`, `bun run docs:build`, `bun run changelog:lint`, `bun run format:check:src`, `bun run typecheck`, `bun run build`, `bun run test:smoke:dist`, and `git diff --check` passed. Earlier full fast/env suites also passed during this task (`bun run test:fast`, `bun run test:env`) before the final N4-only default tweak; focused config/docs/build/smoke checks were rerun after that tweak.

Scope correction from user: secondary subtitles should keep the previous font stack. Interpreting the typo `snas-serif` as the previous literal default `sans-serif`.

Correction verification: restored secondary subtitle default font to `Inter, Noto Sans, Helvetica Neue, sans-serif`. Passed `bun run test:config`, `bun run docs:test`, `bun run changelog:lint`, `bun run format:check:src`, `bun run typecheck`, and `git diff --check`. Verified generated examples and docs contain the restored secondary font.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Updated fresh-install defaults so texthooker startup, the regular subtitle websocket, and the annotation websocket are disabled by default while preserving Yomitan popup auto-pause. Updated primary subtitle defaults to the requested Japanese font stack, kept secondary subtitle font at the prior `Inter, Noto Sans, Helvetica Neue, sans-serif` stack, and applied shared stronger text shadow, transparent backgrounds, and teal N4/fourth-band color. Regenerated config examples, updated docs-site configuration docs, and added a changelog fragment.

Verification passed: `bun run test:config`, `bun run docs:test`, `bun run docs:build`, `bun run changelog:lint`, `bun run format:check:src`, `bun run typecheck`, `bun run build`, `bun run test:smoke:dist`, and `git diff --check`. After the secondary-font correction, reran `bun run test:config`, `bun run docs:test`, `bun run changelog:lint`, `bun run format:check:src`, `bun run typecheck`, and `git diff --check`. `bun run test:fast` and `bun run test:env` also passed during the task before the final small default corrections; affected focused checks were rerun afterward.
<!-- SECTION:FINAL_SUMMARY:END -->
