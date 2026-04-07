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
