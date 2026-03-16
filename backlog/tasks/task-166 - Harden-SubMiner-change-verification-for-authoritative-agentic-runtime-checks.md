---
id: TASK-166
title: Harden SubMiner change verification for authoritative agentic runtime checks
status: Done
assignee: []
created_date: '2026-03-13 05:19'
updated_date: '2026-03-16 05:13'
labels:
  - testing
  - agents
  - verification
dependencies: []
references:
  - >-
    /home/sudacode/projects/japanese/SubMiner/.agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh
  - >-
    /home/sudacode/projects/japanese/SubMiner/.agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh
  - >-
    /home/sudacode/projects/japanese/SubMiner/.agents/skills/subminer-change-verification/SKILL.md
documentation:
  - /home/sudacode/projects/japanese/SubMiner/testing-plan.md
  - /home/sudacode/projects/japanese/SubMiner/docs-site/development.md
ordinal: 22500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Tighten the SubMiner change-verification classifier and verifier so the implementation matches the approved agentic verification plan: authoritative runtime verification must fail closed when unavailable, lane naming should use real-runtime semantics, session and artifact identities must be unique, and the verifier must be safer for parallel agent use.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 The verifier uses `real-runtime` terminology instead of `real-gui` for authoritative runtime verification
- [x] #2 Requested authoritative runtime verification fails closed with a non-green outcome when it cannot run, and unknown lanes do not pass open
- [x] #3 The verifier allocates a unique session identifier and artifact root that does not rely on second-level timestamp uniqueness alone
- [x] #4 The verifier summary/report output includes explicit top-level status and session metadata needed for agent aggregation
- [x] #5 The classifier and verifier better reflect runtime-escalation cases for launcher/plugin/socket/runtime-sensitive changes
- [x] #6 Regression tests cover the new verifier/classifier behavior
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add regression tests for classifier/verifier behavior before changing the scripts.
2. Harden `verify_subminer_change.sh` to use `real-runtime` terminology, fail closed for blocked/unknown authoritative verification, and emit unique session metadata in summaries.
3. Update `classify_subminer_diff.sh` and the skill doc to use `real-runtime` escalation language and better flag launcher/plugin/runtime-sensitive paths.
4. Run targeted regression tests plus a focused dry-run verifier check, then record outcomes and blockers in the task.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added `scripts/subminer-change-verification.test.ts` to regression-test classifier/verifier behavior around `real-runtime` naming, fail-closed authoritative verification, unknown lanes, and unique session metadata.

Reworked `verify_subminer_change.sh` to normalize `real-gui` to `real-runtime`, emit unique session IDs and richer summary metadata, block authoritative runtime verification when unavailable, and fail closed for unknown lanes.

Updated `classify_subminer_diff.sh` to emit `real-runtime-candidate` for launcher/plugin/runtime-sensitive paths, and updated the active skill doc wording to match the new lane terminology.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Hardened the SubMiner change-verification tooling to match the approved agentic verification plan. The verifier now uses `real-runtime` terminology for authoritative runtime verification, preserves compatibility with the deprecated `real-gui` alias, fails closed for unknown lanes, and returns a blocked non-green outcome when requested authoritative runtime verification cannot run. It now allocates a unique session ID and artifact root by default, writes richer session metadata and top-level status into `summary.json`/`summary.txt` plus `reports/summary.*`, and records path selection mode, blockers, and session-local env roots for agent aggregation. The classifier now emits `real-runtime-candidate` for launcher/plugin/runtime-sensitive paths, and the active skill doc uses the same terminology. Verification ran via `bun test scripts/subminer-change-verification.test.ts`, direct dry-run smoke checks for blocked `real-runtime` and failed unknown-lane execution, and a targeted classifier invocation for launcher/plugin paths.
<!-- SECTION:FINAL_SUMMARY:END -->
