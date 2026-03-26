---
id: TASK-159
title: Create SubMiner automated testing skill for agents
status: Done
assignee:
  - codex
created_date: '2026-03-11 05:55'
updated_date: '2026-03-16 05:13'
labels:
  - tooling
  - testing
  - skills
dependencies: []
priority: medium
ordinal: 29500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Design and implement a SubMiner-specific skill that agents can use to verify code changes with automated checks across launcher, mpv, overlay, and related workflows.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Define the skill trigger surface and intended use cases for SubMiner change verification.
- [x] #2 Implement a SubMiner-specific skill package with concise workflow guidance and any required helper scripts or references.
- [x] #3 Document how agents should run automated verification for common SubMiner change types, including launcher/mpv/overlay-sensitive changes.
- [x] #4 Verify the skill itself is usable in this repo context.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a design doc at `docs/plans/2026-03-10-subminer-change-verification-design.md` capturing the approved skill contract, lane selection rules, helper script design, and trigger/workflow guidance.
2. Create an in-repo source package for the new SubMiner-specific verification skill with `SKILL.md` plus helper shell scripts for diff classification and coordinated verification/artifact capture.
3. Implement cheap-first verification logic: map changed paths to verification lanes, run repo-native commands, capture stdout/stderr and metadata under `.tmp/skill-verification/<timestamp>/`, and report pass/fail/skipped states with artifact paths.
4. Encode repo-specific heuristics for docs/config, launcher/plugin, runtime/dist, and optional real-GUI escalation guidance without making GUI launch the default.
5. Verify the skill locally by running the classifier/verifier against representative paths and confirming artifact generation and summaries.
6. Optionally install the finished skill to `~/.codex/skills/subminer-change-verification/` after user approval because that target is outside the writable sandbox.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Planned outputs: SKILL.md plus small helper shell scripts for diff classification and coordinated verification/artifact capture. Default to repo-native commands; only escalate to real GUI/mpv verification when affected paths or requested behavior require it.

2026-03-10: Brainstorming completed with user approval for a SubMiner-specific automated verification skill. Chosen shape: auto-triggering, cheap-first verification workflow with optional real GUI/mpv escalation and artifact capture.

2026-03-10: Added design doc at docs/plans/2026-03-10-subminer-change-verification-design.md and created the in-repo skill source at tools/skills/subminer-change-verification/.

2026-03-10: Verified classifier output with representative launcher/runtime/config/docs paths and ran the verifier in dry-run mode plus one real config lane (`bun run test:config`). Artifacts written under .tmp/skill-verification/20260310-231320 and .tmp/skill-verification/20260310-231326.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented a SubMiner-specific change verification skill source with a concise SKILL.md, a diff classifier, and a cheap-first verifier that captures artifacts under `.tmp/skill-verification/`. Verified the flow with representative path classification, a dry-run multi-lane plan, and a passing real config verification run.
<!-- SECTION:FINAL_SUMMARY:END -->
