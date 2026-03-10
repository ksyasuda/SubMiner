---
id: TASK-134
title: Harden Windows release signing against transient SignPath failures
status: Done
assignee:
  - codex
created_date: '2026-03-09 00:00'
updated_date: '2026-03-08 20:23'
labels:
  - ci
  - release
  - windows
  - signing
dependencies: []
references:
  - .github/workflows/release.yml
  - package.json
  - src/release-workflow.test.ts
  - https://github.com/ksyasuda/SubMiner/actions/runs/22836585479
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

The tag-driven Release workflow currently fails the Windows lane if the SignPath connector returns transient 502 errors during submission, and the tagged build scripts also allow electron-builder to implicitly publish unsigned artifacts before the final release job runs. Harden the workflow so transient SignPath outages get bounded retries and release packaging never auto-publishes unsigned assets.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Windows release signing retries transient SignPath submission failures within the release workflow before failing the job.
- [ ] #2 Release packaging scripts disable electron-builder implicit publish so build jobs do not upload unsigned assets on tag builds.
- [ ] #3 Regression coverage fails if SignPath retry scaffolding or publish suppression is removed.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Add a regression test for the release workflow/package script shape covering SignPath retries and `--publish never`.
2. Patch the Windows release job to retry SignPath submission a bounded number of times and still fail hard if every attempt fails.
3. Update tagged package build scripts to disable implicit electron-builder publishing during release builds.
4. Run targeted release-workflow verification and capture any remaining manual release cleanup needed for `v0.5.0`.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

The failed Windows signing step in GitHub Actions run `22836585479` was not caused by missing secrets or an artifact-shape mismatch. The SignPath GitHub action retried repeated `502` responses from the SignPath connector for several minutes and then failed the job.

Hardened `.github/workflows/release.yml` by replacing the single SignPath submission with three bounded attempts. The second and third submissions only run if the previous attempt failed, and the job now fails with an explicit rerun message only after all three attempts fail. Signed-artifact upload is keyed to the successful attempt so the release job still consumes the normal `windows` artifact name.

Also fixed a separate release regression exposed by the same run: `electron-builder` was implicitly publishing unsigned release assets during tag builds because the packaging scripts did not set `--publish never` and the workflow injected `GH_TOKEN` into build jobs. Updated the relevant package scripts to pass `--publish never`, removed `GH_TOKEN` from the packaging jobs, and made the final publish step force `--draft=false` when editing an existing tag release so previously-created draft releases get published.

Verification: `bun test src/release-workflow.test.ts`, `bun run typecheck`, and `bun run test:fast` all passed locally after restoring the missing local `libsql` install with `bun install --frozen-lockfile`.

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Windows release signing is now resilient to transient SignPath connector outages. The release workflow retries the SignPath submission up to three times before failing, and only uploads the signed Windows artifact from the attempt that succeeded.

Release packaging also no longer auto-publishes unsigned assets on tag builds. The `electron-builder` scripts now force `--publish never`, the build jobs no longer pass `GH_TOKEN` into packaging steps, and the final GitHub release publish step explicitly clears draft state when updating an existing tag release.

Validation: `bun test src/release-workflow.test.ts`, `bun run typecheck`, `bun run test:fast`.
Manual follow-up for the failed `v0.5.0` release: rerun the `Release` workflow after merging/pushing this fix, then clean up the stray draft/untagged release assets created by the failed run if they remain.

<!-- SECTION:FINAL_SUMMARY:END -->
