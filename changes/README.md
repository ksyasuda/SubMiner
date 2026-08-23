# Changelog Fragments

Add one `.md` file per user-visible PR in this directory.

Use this format:

```md
type: added
area: overlay

- Added keyboard navigation for Yomitan popups.
- Added auto-pause toggle when opening the popup.
```

For breaking changes, add `breaking: true`:

```md
type: changed
area: config
breaking: true

- Renamed `foo.bar` to `foo.baz`.
```

Rules:

- `type` required: `added`, `changed`, `fixed`, `docs`, or `internal`
- `area` required: short product area like `overlay`, `launcher`, `release`
- `breaking` optional: set to `true` to flag as a breaking change
- each non-empty body line becomes a bullet
- `README.md` is ignored by the generator
- if a PR should not produce release notes, apply the `skip-changelog` label instead of adding a fragment

PR branch workflow:

- Before adding a fragment or bullet, inspect the `changes/*.md` files already changed in the PR
- If the new work fixes, modifies, renames, or supersedes behavior introduced or referenced by that fragment, edit or remove the stale bullet instead of adding follow-up churn
- Add a new bullet only when it describes a truly separate user-visible outcome
- Multiple fragment files are allowed when one PR has genuinely separate release-note outcomes, but keep them minimized and current

How fragments turn into a release:

- At release time, `bun run changelog:build` (and `bun run changelog:prerelease-notes`) pipes every pending fragment through `claude -p` to merge related items, drop noise, and rewrite into a clean user-facing release body. Write fragments as raw, informative notes — don't worry about polished prose, deduping across PRs, or line-by-line phrasing. The polish step handles all of that.
- The polish step treats pending fragments as the final release outcome, not prerelease history. If a feature is added and then renamed or fixed before the stable cut, ship the final feature bullet instead of separate prerelease-only breaking/fix entries.
- GitHub release notes and prerelease notes use short top-level items with nested bullets for the change, user benefit, and any useful action note. The stable `CHANGELOG.md` can stay in compact single-line bullets.
- `internal` fragments stay in `CHANGELOG.md` (inside a collapsed `<details>` block) but are dropped from the GitHub release notes entirely.
- The polished `CHANGELOG.md` and `release/release-notes.md` are committed and reviewed before tagging — edit the Markdown by hand if Claude misses something.

Prerelease notes:

- prerelease tags like `v0.11.3-beta.1` and `v0.11.3-rc.1` reuse the current pending fragments to generate `release/prerelease-notes.md`
- from the second prerelease of a base version onward, the notes also open with a `## Changes since <previous tag>` section generated from the fragment diff against the previous beta/RC tag; keep fragment edits meaningful — editorial-only rewording is filtered out of that section, while genuinely changed behavior and deleted fragments (reverted changes) are reported
- existing prerelease notes are a reviewed baseline; later prerelease runs should replace stale beta/RC wording with the current outcome instead of appending fix churn
- prerelease note generation does not consume fragments and does not update `CHANGELOG.md` or `docs-site/changelog.md`
- the final stable release is the point where `bun run changelog:build` consumes fragments into the stable changelog and release notes
