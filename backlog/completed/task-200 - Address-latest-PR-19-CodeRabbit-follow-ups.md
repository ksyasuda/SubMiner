---
id: TASK-200
title: 'Address latest PR #19 CodeRabbit follow-ups'
status: Done
assignee:
  - '@codex'
created_date: '2026-03-19 07:18'
updated_date: '2026-03-23 03:22'
labels:
  - pr-review
  - anki-integration
  - launcher
milestone: m-1
dependencies: []
references:
  - launcher/mpv.test.ts
  - src/anki-integration.ts
  - src/anki-integration/card-creation.ts
  - src/anki-integration/runtime.ts
  - src/anki-integration/known-word-cache.ts
priority: medium
ordinal: 133500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Validate the latest 2026-03-19 CodeRabbit review round on PR #19, implement only the confirmed fixes, and verify the touched launcher and Anki integration paths.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Each latest-round PR #19 CodeRabbit inline comment is validated against the current branch and classified as actionable or not warranted
- [x] #2 Confirmed correctness issues in launcher and Anki integration code are fixed with focused regression coverage where practical
- [x] #3 Targeted verification runs for the touched areas and the task notes record what changed versus what was rejected
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Validate the five inline comments from the 2026-03-19 CodeRabbit PR #19 review against current launcher and Anki integration code.
2. Add or extend focused tests for any confirmed launcher env-sandbox, notification-state, AVIF lead-in propagation, or known-word-cache lifecycle/scope regressions.
3. Apply the smallest safe fixes in `launcher/mpv.test.ts`, `src/anki-integration.ts`, `src/anki-integration/card-creation.ts`, `src/anki-integration/runtime.ts`, and `src/anki-integration/known-word-cache.ts` as needed.
4. Run targeted unit tests plus the SubMiner verification helper on the touched files, then record which comments were accepted or rejected in task notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Validated the five latest inline comments from CodeRabbit review `3973222927` on PR #19.

Accepted fixes:
- Hardened the three `findAppBinary` launcher tests against host leakage by sandboxing `SUBMINER_APPIMAGE_PATH` / `SUBMINER_BINARY_PATH` and stubbing executable checks so `/opt` and PATH resolution are deterministic.
- `showNotification()` now marks OSD/both updates as failed when `errorSuffix` is present instead of always rendering a success marker.
- `applyRuntimeConfigPatch()` now avoids starting or stopping known-word cache lifecycle work while the runtime is stopped, while still clearing cached state when highlighting is disabled.
- Extracted shared known-word cache lifecycle helpers and switched the persisted cache identity to the same lifecycle config used by runtime restart detection, so changes to `fields.word`, per-deck field mappings, or refresh interval invalidate stale cache state correctly.

Rejected fix:
- The `createSentenceCard()` AVIF lead-in comment was technically incomplete for this branch. There is no current caller that computes an `animatedLeadInSeconds` input for sentence-card creation, and the existing lead-in resolver depends on note media fields that do not exist before the new card's media is generated.

Regression coverage added:
- `src/anki-integration.test.ts` partial-failure OSD result marker.
- `src/anki-integration/runtime.test.ts` stopped-runtime known-word lifecycle guards.
- `src/anki-integration/known-word-cache.test.ts` cache invalidation when `fields.word` or per-deck field mappings change.

Verification:
- `bun test src/anki-integration/runtime.test.ts`
- `bun test src/anki-integration/known-word-cache.test.ts`
- `bun test src/anki-integration.test.ts --test-name-pattern 'marks partial update notifications as failures in OSD mode'`
- `bun test launcher/mpv.test.ts --test-name-pattern 'findAppBinary resolves ~/.local/bin/SubMiner.AppImage when it exists|findAppBinary resolves /opt/SubMiner/SubMiner.AppImage when ~/.local/bin candidate does not exist|findAppBinary finds subminer on PATH when AppImage candidates do not exist'`
- `bun test src/anki-integration.test.ts src/anki-integration/runtime.test.ts src/anki-integration/known-word-cache.test.ts launcher/mpv.test.ts`
- `bash .agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh launcher/mpv.test.ts src/anki-integration.ts src/anki-integration/runtime.ts src/anki-integration/known-word-cache.ts src/anki-integration/runtime.test.ts src/anki-integration/known-word-cache.test.ts src/anki-integration.test.ts`
- `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane launcher-plugin --lane core launcher/mpv.test.ts src/anki-integration.ts src/anki-integration/runtime.ts src/anki-integration/known-word-cache.ts src/anki-integration/runtime.test.ts src/anki-integration/known-word-cache.test.ts src/anki-integration.test.ts`

Verifier result:
- `launcher-plugin` lane passed (`test:launcher:smoke:src`, `test:plugin:src`).
- `core/typecheck` passed.
- `core/test-fast` failed for an unrelated existing environment issue in `scripts/update-aur-package.test.ts`: `scripts/update-aur-package.sh: line 71: mapfile: command not found` under the local macOS Bash environment.
- Verifier artifacts: `.tmp/skill-verification/subminer-verify-20260319-002617-UgpKUy`

Classification: actionable and fixed -> `launcher/mpv.test.ts` env leakage hardening, `src/anki-integration.ts` partial-failure OSD marker, `src/anki-integration/runtime.ts` started-guard for known-word lifecycle calls, `src/anki-integration/known-word-cache.ts` cache identity alignment with runtime lifecycle config.

Classification: not warranted as written -> `src/anki-integration/card-creation.ts` lead-in threading comment. No current `createSentenceCard()` caller computes or owns an `animatedLeadInSeconds` value, and the existing lead-in helper derives from preexisting note media fields, so blindly adding an optional parameter would not fix a real branch behavior bug.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed four confirmed PR #19 latest-round CodeRabbit issues locally: deterministic launcher `findAppBinary` tests, correct partial-failure OSD result markers, started-state guards around known-word cache lifecycle restarts, and shared known-word cache identity logic so field-mapping changes invalidate stale cache state. Added focused regression coverage for each confirmed behavior.

One comment was intentionally not applied: the `createSentenceCard()` AVIF lead-in suggestion does not match the current branch architecture because no caller computes that value today and the existing resolver requires preexisting note media fields. Verification is green for all touched targeted tests plus the launcher-plugin/core typecheck lanes; the only remaining red is an unrelated existing `test:fast` failure in `scripts/update-aur-package.test.ts` caused by `mapfile` being unavailable in the local Bash environment.
<!-- SECTION:FINAL_SUMMARY:END -->
