# AI-Polished Changelog Design

Date: 2026-05-02
Status: Approved, awaiting implementation plan

## Problem

Today every user-visible PR drops a fragment under `changes/`. The
fragments are written by whoever shipped the PR, so the prose tone,
granularity, and noise level vary wildly. By the time a stable release
is cut, the assembled CHANGELOG section can run 40+ bullets and lean
heavily on internal area labels (`Overlay:`, `Launcher:`, etc.). End
users skim it, miss the actual story of the release, and the GitHub
release body inherits the same clutter.

## Goal

Replace the mechanical `renderGroupedChanges` step in
`scripts/build-changelog.ts` with a single non-interactive `claude -p`
call that:

- Merges related fragments into coherent feature-level bullets.
- Drops PR-housekeeping noise that no user benefits from reading.
- Reorganizes by user-visible feature, not by `area:` prefix.
- Produces a tight `## v<version>` body suitable for both
  `CHANGELOG.md` and the GitHub release notes.

The fragment file format and authoring workflow stay the same. Only
the rendering step changes.

## Decisions

1. **Pipeline:** Replace the current bullet renderer. `changelog:build`
   and `changelog:prerelease-notes` always invoke `claude -p`. There
   is no legacy fallback path.
2. **Review:** Write straight to `CHANGELOG.md` /
   `release/release-notes.md` / `release/prerelease-notes.md`. The
   release engineer reviews the diff before tagging; no extra
   confirmation prompt.
3. **CI:** Local-only. The release workflow no longer auto-runs
   `changelog:build` when fragments are pending; instead it fails the
   tagged release with a clear message asking the user to run the
   build locally and commit the result. CI does not need
   `claude` on PATH or any Anthropic credentials.
4. **Polish depth:** Heavy rewrite plus dedupe. Claude is allowed to
   merge bullets, drop trivial ones, and reorder by feature.
5. **Internal section:** Dropped from release notes entirely. Kept in
   `CHANGELOG.md` inside a `<details><summary>Internal changes</summary>`
   collapse so the historical record survives without taking over the
   page.
6. **Prereleases:** Same polish path. Beta and RC notes look as clean
   as stable.
7. **Claude invocation:**
   `claude -p --bare --model sonnet --permission-mode bypassPermissions --output-format text`
   over stdin. `--bare` skips hooks, MCP, auto-memory, and CLAUDE.md
   discovery so the call is self-contained and reproducible.
8. **Failure mode:** Hard fail. Missing `claude` binary, non-zero
   exit, empty/short output, or output missing the required section
   headers all abort the build.

## Architecture

### Affected files

- `scripts/build-changelog.ts` — primary edit. Replace
  `renderGroupedChanges` with `polishFragmentsWithClaude`. Wire it
  into `writeChangelogArtifacts`, `writePrereleaseNotesForVersion`,
  and any other path that currently calls `renderGroupedChanges`.
- `scripts/build-changelog.test.ts` — add `runClaude` to the dep
  injection surface, stub it in unit tests, add new coverage for
  failure modes and mode='changelog' vs mode='release-notes'.
- `.github/workflows/release.yml` (or wherever the auto-build lives)
  — remove the auto-`changelog:build` fallback; add a guard that
  fails the run if `changes/*.md` exist on a tag-based release.
- `src/release-workflow.test.ts`,
  `src/prerelease-workflow.test.ts` — update if they exercise the
  CI auto-run path.
- `docs/RELEASING.md` — document the new local-only build step and
  the `claude` PATH requirement.
- `changes/README.md` — note that fragments will be merged and
  rewritten, so authors should write raw notes rather than polished
  prose.
- `changes/<n>-ai-changelog-polish.md` — fragment for this change
  itself (`type: internal`, `area: release`).

### New function: `polishFragmentsWithClaude`

```
polishFragmentsWithClaude(
  fragments: ChangeFragment[],
  options: {
    mode: 'changelog' | 'release-notes',
    version: string,
    date?: string,        // changelog mode only
    deps?: { runClaude?: (prompt: string) => string },
  },
): string
```

- Filters out `internal` fragments when `mode === 'release-notes'`.
- Serializes the surviving fragments into the prompt format below.
- Invokes `claude` (via `deps.runClaude` or the default
  `execFileSync` wrapper).
- Validates the output contains the expected section headers; throws
  on failure.
- Returns the Markdown body that will be inserted under the
  `## v<version>` heading.

The existing `prependReleaseSection`, `extractReleaseSectionBody`,
`writeReleaseNotesFile`, and `generateDocsChangelog` functions stay
exactly as they are — they consume the polished body the same way
they consumed the rendered body.

### Prompt input format

```
MODE: changelog | release-notes
VERSION: 0.13.0
DATE: 2026-05-02            (changelog mode only)

FRAGMENT changes/291-character-dictionary-selection.md
type: added
area: dictionary
breaking: false
- Added CLI and in-app AniList selection for character dictionary mismatches...
- Added launcher support through `subminer dictionary --candidates`...

FRAGMENT changes/293-interjection-annotation-filter.md
type: fixed
area: tokenizer
breaking: false
- Stopped standalone `あ` interjections from receiving subtitle annotation metadata...

...
```

### Prompt output contract

Claude is instructed to emit Markdown only, no preamble, conforming to:

- Sections in this order, omitting empty ones:
  `### Breaking Changes`, `### Added`, `### Changed`,
  `### Fixed`, `### Docs`.
- For `mode === 'changelog'`, an additional terminal section:

  ```
  <details>
  <summary>Internal changes</summary>

  ### Internal
  - …
  - …

  </details>
  ```

- Bullets lead with a feature name (e.g. `Playlist browser:`,
  `Windows overlay:`) — not with the raw `area:` slug.
- Bullets may merge multiple fragments, but breaking changes must
  retain their substance and stay in `### Breaking Changes`.
- Bullets that document only PR-housekeeping or CodeRabbit follow-ups
  are dropped unless they describe a user-visible behavior change.

After parsing, the body is passed unchanged to `prependReleaseSection`,
which wraps it under the standard `## v<version> (<date>)` heading.

### Determinism and review

Claude is non-deterministic. The mitigation is the existing release
workflow: the polished CHANGELOG/release-notes are committed before
tagging, so what's in the repo is what ships. The release engineer
diffs the commit before pushing.

## Testing

`scripts/build-changelog.test.ts` already exposes a `deps` injection
seam for fs operations. Extend it with `runClaude?: (prompt: string)
=> string`. Tests stub it to return canned Markdown; the default
implementation (used in production) wraps `execFileSync('claude', …)`.

New cases to cover:

- Mode `'release-notes'` filters `internal` fragments before sending
  to Claude.
- Mode `'changelog'` includes `internal` fragments and the
  `<details>` wrapper expectation is propagated through to the prompt.
- Missing `claude` binary throws a clear error.
- Non-zero exit code from `claude` throws.
- Empty / whitespace-only output throws.
- Output missing the expected section headers throws.
- Prerelease path uses release-notes mode and writes
  `release/prerelease-notes.md` with the existing disclaimer.
- The polished body still flows correctly through
  `prependReleaseSection`, `extractReleaseSectionBody`, and
  `generateDocsChangelog`.

`release-workflow.test.ts` and `prerelease-workflow.test.ts` are
updated to cover the new "fragments pending on a tag" failure mode.

## Documentation

- `docs/RELEASING.md` updates the stable and prerelease sections to
  state that `changelog:build` and `changelog:prerelease-notes` invoke
  `claude` locally, list the PATH requirement, and remove the note
  about CI auto-running `changelog:build`.
- `changes/README.md` adds a short note: fragments will be merged and
  rewritten by Claude during release, so write raw notes — don't
  polish them.

## Out of Scope (YAGNI)

- No caching of polished output between runs.
- No diff-and-confirm interactive prompt.
- No SDK fallback if the CLI is missing.
- No `--no-polish` escape hatch. Revert the commit if you genuinely
  need the legacy renderer back.
- No change to the fragment authoring schema or the `pr-check`
  enforcement.

## Open Risks

- **Claude regresses on a release.** Mitigated by the
  commit-before-tag review step. If the polish is bad, the release
  engineer edits `CHANGELOG.md` directly and re-commits before
  tagging.
- **`--bare` mode behavior changes upstream.** Pinning to `--model
  sonnet` and `--bare` reduces drift, but Claude Code is still an
  external tool. If invocation flags change, the release engineer
  will see a hard failure and can patch the script.
- **CI surface area increases.** Removing the auto-`changelog:build`
  fallback means a tag pushed before running the local build will
  fail CI. The error message must clearly point at the fix.
