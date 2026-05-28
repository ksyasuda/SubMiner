type: internal
area: release

- Release-note polishing treats pending fragments and reviewed prerelease notes as a cumulative final outcome, collapsing prerelease-only fixes or breakages into the final user-facing change.
- Prerelease note generation reuses existing reviewed notes and merges only new fragment material, and `make clean` preserves `release/prerelease-notes.md`.
- Release-note polishing now asks Claude to write short, nested highlight bullets so longer changes are easier to scan.
