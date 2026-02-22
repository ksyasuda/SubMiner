---
id: TASK-82
title: Add end-to-end smoke suite for launcher mpv ipc and overlay runtime
status: Done
assignee:
  - codex-task82-smoke-20260222T002523Z-3j7u
created_date: '2026-02-18 11:43'
updated_date: '2026-02-22 00:55'
labels:
  - testing
  - e2e
  - launcher
  - mpv
dependencies:
  - TASK-74
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Current coverage is strong at unit/service level but thin on end-to-end behavior across launcher, mpv socket lifecycle, IPC wiring, and overlay runtime initialization. This task introduces an observable e2e smoke suite to reduce refactor risk.
<!-- SECTION:DESCRIPTION:END -->

## Suggestions

<!-- SECTION:SUGGESTIONS:BEGIN -->
- Start with deterministic smoke flows, not full UI automation.
- Use minimal fixtures and fakeable mpv socket harness where possible.
- Treat this as release gate signal, with clear pass/fail logs.
<!-- SECTION:SUGGESTIONS:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Define smoke scenarios:
   - launcher command branch sanity
   - mpv socket ready/not-ready handling
   - IPC channel registration + key command dispatch
   - overlay runtime bootstrap lifecycle
2. Build test harness that can run headless in CI/local environments.
3. Add stable fixtures/mocks for mpv and external tool boundaries.
4. Produce concise logs/artifacts for failures.
5. Integrate smoke suite into pre-release validation flow.
6. Document smoke test usage and troubleshooting.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 E2E smoke suite exists and runs in automated workflow
- [x] #2 Core launcher/mpv/ipc/overlay startup scenarios are covered
- [x] #3 Failures produce actionable logs/artifacts
- [x] #4 Smoke suite documented in development/release docs
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Plan-of-record (2026-02-22, codex-task82-smoke-20260222T002523Z-3j7u)

Goal
- Add deterministic e2e smoke coverage for launcher -> mpv IPC socket -> app overlay start/stop runtime wiring.
- Run suite in CI + release quality gates with actionable failure artifacts.
- Document command, coverage, and release-checklist linkage.

Execution slices (parallel where safe)
1) Launcher smoke harness (core implementation)
   - Create launcher/smoke.e2e.test.ts with fake mpv + fake app fixtures in temp dir.
   - Validate startup path, socket readiness, app --start args/env, and post-run --stop behavior.
   - Emit actionable artifact logs under .tmp/launcher-smoke/<run-id>.

2) Script + workflow wiring (can run in parallel after command exists)
   - package.json: add test:launcher:smoke:src and include in launcher lane.
   - .github/workflows/ci.yml: run launcher smoke suite and upload smoke artifacts on failure.
   - .github/workflows/release.yml: mirror quality-gate step and failure artifact upload.

3) Docs wiring (parallel with workflows)
   - docs/development.md: command usage + what it validates + artifact path.
   - docs/installation.md: release/checklist reference to launcher smoke gate.

4) Validation + closure
   - Run focused lane: bun run test:launcher:smoke:src and bun run test:launcher.
   - Run integration lane: bun run test:fast.
   - Run dist confidence lane: bun run build && bun run test:smoke:dist (or record unrelated blockers).
   - Update TASK-82 AC/DoD checks with exact evidence and final summary.

Scope guardrails
- No commit/push in this run.
- Keep smoke suite dependency-free (local fakes only).
- If new required scope appears, pause and request user scope decision before expanding AC.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-02-22: Started execution with opencode-task82-smoke-20260222T002150Z-p5bp via writing-plans -> executing-plans workflow.

Implemented launcher e2e smoke suite in `launcher/smoke.e2e.test.ts` using local fake `mpv` + fake app binaries, temp config/runtime sockets, and assertions for mpv status readiness, launcher startup wiring (`--backend`, `--socket`), IPC socket forwarding, overlay env propagation (`SUBMINER_MPV_LOG`), and post-mpv overlay stop command.

Added automation wiring: `package.json` scripts (`test:launcher:smoke:src`) and CI/release quality-gate steps to run smoke suite plus upload `.tmp/launcher-smoke/**` artifacts on failure in `.github/workflows/ci.yml` and `.github/workflows/release.yml`.

Documented smoke workflow and artifact triage in `docs/development.md` and release/pre-release verification guidance in `docs/installation.md`.

Verification evidence: `bun run test:launcher:smoke:src` PASS, `bun run test:launcher` PASS, `bun run test:fast` PASS, `bun run docs:build` PASS. `bun run build && bun run test:smoke:dist` remains blocked by pre-existing unrelated TASK-80 IPC typing errors in `src/core/services/ipc.ts` + `src/core/services/ipc.test.ts`.

2026-02-22: Implemented launcher smoke suite at `launcher/smoke.e2e.test.ts` with deterministic fake mpv/app harness and preserved failure artifacts under `.tmp/launcher-smoke/<case-id>`.

2026-02-22: Wired source smoke lane into automation: `package.json` (`test:launcher:smoke:src`), `.github/workflows/ci.yml` and `.github/workflows/release.yml` now run the lane and upload `.tmp/launcher-smoke/**` artifacts on failure.

2026-02-22: Updated docs with smoke command + release gate reference and artifact guidance in `docs/development.md` and `docs/installation.md`.

2026-02-22 verification: `bun run test:launcher:smoke:src`, `bun run test:launcher`, `bun run test:fast`, `bun run build && bun run test:smoke:dist`, `bun run docs:build`.

2026-02-22 correction: `bun run build && bun run test:smoke:dist` passes in current workspace; prior note about TASK-80-related blockage is stale from earlier intermediate run.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added an end-to-end launcher smoke suite that exercises launcher->mpv IPC socket lifecycle plus overlay start/stop runtime wiring using deterministic local fake mpv/app fixtures and artifact-preserving failure output. Integrated the smoke lane into CI and release quality gates with failure artifact uploads, documented local/release usage, and verified all required test/build/docs lanes pass.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Smoke suite passes on baseline branch
- [x] #2 Release checklist references smoke suite
<!-- DOD:END -->
