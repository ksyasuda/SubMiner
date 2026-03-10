<!-- read_when: cutting a tagged release or debugging release prep -->

# Releasing

1. Confirm `main` is green: `gh run list --workflow CI --limit 5`.
2. Bump `package.json` to the release version.
3. Build release metadata before tagging:
   `bun run changelog:build --version <version>`
4. Review `CHANGELOG.md`.
5. Run release gate locally:
   `bun run changelog:check --version <version>`
   `bun run test:fast`
   `bun run typecheck`
6. Commit release prep.
7. Tag the commit: `git tag v<version>`.
8. Push commit + tag.

Notes:

- `changelog:check` now rejects tag/package version mismatches.
- Do not tag while `changes/*.md` fragments still exist.
