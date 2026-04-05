---
id: TASK-284
title: Fix CI changelog fragment requirement for launcher X11 MPV fallback change
status: Done
assignee: []
created_date: '2026-04-05 22:21'
updated_date: '2026-04-05 22:22'
labels:
  - ci
  - changelog
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Current CI is failing in `changelog:pr-check` because PR changes release-relevant files but does not include a required entry under `changes/` and lacks `skip-changelog` label. Add a release fragment describing the behavioral change and verify the CI gate passes.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Add a correctly formatted changelog fragment under `changes/` for the current change
- [x] #2 Run the local changelog PR check (or equivalent) with passing result
- [x] #3 Run required CI gate commands after change and confirm no regressions
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added `changes/2026.04.05-mpv-x11-fallback.md` with launcher release-note metadata so `changelog:pr-check` can pass. Verified local CI gate commands: `bun run typecheck`, `bun run test:fast`, `bun run test:env`, `bun run build`, `bun run test:smoke:dist` all passed. Ran manual PR changelog verification by invoking `verifyPullRequestChangelog` with current git diff plus the new fragment and confirmed it passes.
<!-- SECTION:FINAL_SUMMARY:END -->
