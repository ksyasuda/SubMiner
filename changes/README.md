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

How fragments turn into a release:

- At release time, `bun run changelog:build` (and `bun run changelog:prerelease-notes`) pipes every pending fragment through `claude -p` to merge related items, drop noise, and rewrite into a clean user-facing release body. Write fragments as raw, informative notes — don't worry about polished prose, deduping across PRs, or line-by-line phrasing. The polish step handles all of that.
- `internal` fragments stay in `CHANGELOG.md` (inside a collapsed `<details>` block) but are dropped from the GitHub release notes entirely.
- The polished `CHANGELOG.md` and `release/release-notes.md` are committed and reviewed before tagging — edit the Markdown by hand if Claude misses something.

Prerelease notes:

- prerelease tags like `v0.11.3-beta.1` and `v0.11.3-rc.1` reuse the current pending fragments to generate `release/prerelease-notes.md`
- prerelease note generation does not consume fragments and does not update `CHANGELOG.md` or `docs-site/changelog.md`
- the final stable release is the point where `bun run changelog:build` consumes fragments into the stable changelog and release notes
