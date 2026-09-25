---
name: subminer-release
description: Prepare, cut, publish, or repair SubMiner stable and prerelease releases. Use for hands-on release work; do not use for general release questions.
---

# SubMiner release

Carry out the requested release phase using the repository's current release process.

## Source of truth

Read `docs/RELEASING.md` completely before changing files or release state. Treat it as canonical. Read `changes/README.md` when the work touches change fragments or generated release notes.

Do not copy release commands or policy into this skill. If this skill disagrees with the release guide, follow the guide and reconcile the skill before handoff.

## Workflow

1. Identify whether the request is for a stable release, prerelease, release preparation, publication, or repair.
2. Inspect the current branch, worktree status, package version, pending change fragments, relevant tags, and latest CI state before making changes.
3. Follow the matching procedure in `docs/RELEASING.md` in order. Review generated changelog and release-note Markdown before it can be committed or published.
4. Run every required gate for the requested release phase. Do not treat a cheaper test lane as a substitute for the documented release gate.
5. Before a stable tag, confirm the package and tag versions match and no pending `changes/*.md` fragments remain. Preserve fragments for prereleases as documented.
6. Report the resulting version, completed checks, local commit and tag state, remote publication state, skipped platform checks, and any remaining manual work.

## Authorization boundaries

- A request to prepare a release stops before commit, tag, push, or remote publication unless the user also authorizes those actions.
- A clear request to cut or publish a release includes the documented commit, tag, and push steps. Ask before the first remote mutation when the wording is ambiguous.
- Do not edit an existing GitHub release, publish to the AUR, change secrets, or alter signing configuration unless the user explicitly requests that operation.
- Do not switch branches without consent.

## Stop conditions

Stop and report the blocker when required CI or a release gate fails, authentication is missing, versions disagree, required artifacts are absent, or the worktree contains unexpected changes that overlap the release. Do not tag or publish a partially verified release.
