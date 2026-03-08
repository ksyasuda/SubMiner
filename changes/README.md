# Changelog Fragments

Add one `.md` file per user-visible PR in this directory.

Use this format:

```md
type: added
area: overlay

- Added keyboard navigation for Yomitan popups.
- Added auto-pause toggle when opening the popup.
```

Rules:

- `type` required: `added`, `changed`, `fixed`, `docs`, or `internal`
- `area` required: short product area like `overlay`, `launcher`, `release`
- each non-empty body line becomes a bullet
- `README.md` is ignored by the generator
- if a PR should not produce release notes, apply the `skip-changelog` label instead of adding a fragment
