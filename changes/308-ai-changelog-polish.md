type: internal
area: release

- Replaced the changelog renderer with a `claude -p` polish pass that merges related fragments, drops PR housekeeping, and writes user-friendly release notes. CHANGELOG.md keeps internal items in a collapsed `<details>` block; the GitHub release notes drop them entirely.
- Removed the release CI auto-build for pending `changes/*.md` fragments. Tag-based release runs now fail fast with a clear error if fragments are still pending; build the changelog locally with `bun run changelog:build` (which requires the `claude` CLI on PATH) and commit before tagging.
