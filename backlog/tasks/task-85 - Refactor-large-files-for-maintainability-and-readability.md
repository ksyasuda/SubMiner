---
id: TASK-85
title: Refactor large files for maintainability and readability
status: In Progress
assignee: []
created_date: '2026-02-19 09:46'
updated_date: '2026-02-20 11:42'
labels:
  - architecture
  - refactor
  - maintainability
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Several core files are oversized and high-coupling (`src/main.ts`, `src/anki-integration.ts`, `src/config/service.ts`, `src/core/services/immersion-tracker-service.ts`). This task tracks phased, behavior-preserving decomposition plus guardrails and generated-launcher workflow cleanup.
<!-- SECTION:DESCRIPTION:END -->

## Suggestions

<!-- SECTION:SUGGESTIONS:BEGIN -->
- Use seam tests before each extraction.
- Keep `src/main.ts` + `src/anki-integration.ts` as thin composition/coordinator layers.
- Formalize `subminer` as generated artifact only.
<!-- SECTION:SUGGESTIONS:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Add file-budget guardrails and baseline report.
2. Split `src/main.ts` into runtime domain modules (ongoing; major wiring extracted).
3. Split `src/anki-integration.ts` into focused collaborators (started; constructor composition decomposed).
4. Split `src/config/service.ts` by load/migrate/validate/warn phases (started; load/apply flow decomposition in place).
5. Split immersion tracker service by state, persistence, sync responsibilities.
6. Clarify generated launcher artifact workflow and docs.
7. Run full build/test gate and publish maintainability report.
<!-- SECTION:PLAN:END -->

## Progress Notes

<!-- SECTION:PROGRESS:BEGIN -->
- 2026-02-20: Large `src/main.ts` composition slices extracted into runtime handler modules (startup, CLI, tray/window/bootstrap, shortcuts, OSD/secondary-sub, numeric/overlay shortcut lifecycle, and related seams) with focused parity tests.
- 2026-02-20: `src/anki-integration.ts` constructor decomposed into targeted private factory/config methods (`normalizeConfig`, `createKnownWordCache`, `createPollingRunner`, `createCardCreationService`, `createFieldGroupingService`) to reduce orchestration complexity.
- 2026-02-20: `src/config/service.ts` load/apply duplication reduced by introducing `applyResolvedConfig`, `resolveExistingConfigPath`, and `parseConfigContent`.
- Validation checkpoint: `bun run build` and focused suites (`dist/config/config.test.js`, `dist/anki-integration.test.js`) passing.
- Current strategy: prioritize high-ROI extractions (behavioral seams and churn-heavy hotspots) over further low-impact micro-shuffles.
<!-- SECTION:PROGRESS:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 `src/main.ts` reduced to orchestration-focused module with extracted runtime domains
- [ ] #2 `src/anki-integration.ts` reduced to facade with helper collaborators
- [ ] #3 Config and immersion tracker services decomposed without behavior regressions
- [ ] #4 `subminer` generated artifact ownership/workflow documented and enforced
- [ ] #5 Full build + config/core tests pass after refactor
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Plan at `docs/plans/2026-02-19-repo-maintainability-refactor-plan.md` executed or decomposed into child tasks
- [ ] #2 Regression coverage added for extracted seams
- [ ] #3 Docs updated for architecture and contributor workflow changes
<!-- DOD:END -->
