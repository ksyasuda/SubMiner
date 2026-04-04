---
id: TASK-275
title: Patch high-severity audit findings with minimal dependency changes
status: Done
assignee:
  - codex
created_date: '2026-04-04 04:45'
updated_date: '2026-04-04 04:50'
labels:
  - security
  - dependencies
dependencies: []
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Update SubMiner's direct Electron runtime and vulnerable build-time transitive dependencies to patched versions using the smallest safe version moves. Keep electron-builder on the current pinned line unless verification shows a blocker. Verify that bun audit no longer reports the current high-severity findings and that the standard project gate still passes.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Electron is updated to a patched supported release on the current supported line with no broader dependency refresh
- [x] #2 Vulnerable transitive packages @xmldom/xmldom, lodash, and picomatch resolve to patched versions via targeted dependency changes
- [x] #3 `bun audit --audit-level high` no longer reports the currently listed high-severity findings
- [x] #4 The default handoff verification gate passes, or any failure is documented with the exact command and error output
- [x] #5 Any dependency or lockfile changes remain minimal and do not change the pinned electron-builder line unless required
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Update package.json with the smallest set of dependency changes: bump electron from ^37.10.3 to 39.8.6 and add overrides for @xmldom/xmldom 0.8.12, lodash 4.18.0, and picomatch 4.0.4 while leaving electron-builder pinned at 26.8.2.
2. Refresh bun.lock with a lockfile-only install/update and confirm the resolved versions for electron, @xmldom/xmldom, lodash, and picomatch.
3. Run bun audit --audit-level high and verify the current high-severity findings are gone.
4. Run the default verification gate: bun run typecheck, bun run test:fast, bun run test:env, bun run build, bun run test:smoke:dist.
5. If any verification step fails, capture the exact failing command and error, assess whether it is caused by the dependency updates, and stop without broadening scope.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Updated package.json to pin electron 39.8.6 and add overrides for @xmldom/xmldom 0.8.12, lodash 4.18.0, and picomatch 4.0.4 while keeping electron-builder pinned at 26.8.2.

Refreshed bun.lock with bun install and confirmed the patched versions resolved in the lockfile.

Verification passed: bun audit --audit-level high, bun run typecheck, bun run test:fast, bun run test:env, bun run build, bun run test:smoke:dist.

Added changelog fragment changes/patch-audit-dependencies.md for the security/dependency maintenance update. No internal docs or docs-site updates were needed because the change does not alter user-facing behavior, configuration, or workflows.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Cleared the reported high-severity audit findings with minimal dependency churn by pinning electron to 39.8.6 and overriding @xmldom/xmldom, lodash, and picomatch to patched versions. Kept electron-builder on 26.8.2. bun audit is clean and the full default handoff gate passed: typecheck, fast tests, env tests, build, and dist smoke tests.
<!-- SECTION:FINAL_SUMMARY:END -->
