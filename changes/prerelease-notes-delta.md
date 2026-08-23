type: changed
area: release

- Prerelease notes now open with a "Changes since" section that lists only what changed compared to the previous beta/RC of the same version, above the cumulative highlights.
- CI now rejects prerelease tags whose committed notes were generated for a different beta/RC, instead of silently shipping stale notes.
