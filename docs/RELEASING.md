<!-- read_when: cutting a tagged release or debugging release prep -->

# Releasing

## Stable Release

1. Confirm `main` is green: `gh run list --workflow CI --limit 5`.
2. Confirm release-facing docs are current: `README.md`, `changes/*.md`, and any touched `docs-site/` pages/config examples.
3. Run `bun run changelog:lint`.
4. Bump `package.json` to the release version.
5. Build release metadata before tagging:
   `bun run changelog:build --version <version> --date <yyyy-mm-dd>`
   - Release CI now also auto-runs this step when releasing directly from a tag and `changes/*.md` fragments remain.
6. Review `CHANGELOG.md` and `release/release-notes.md`.
7. Run release gate locally:
   `bun run changelog:check --version <version>`
   `bun run verify:config-example`
   `bun run typecheck`
   `bun run test:fast`
   `bun run test:env`
   `bun run build`
8. If `docs-site/` changed, also run:
   `bun run docs:test`
   `bun run docs:build`
9. Commit release prep.
10. Tag the commit: `git tag v<version>`.
11. Push commit + tag.

## Prerelease

1. Confirm release-facing docs and pending `changes/*.md` fragments are current.
2. Run `bun run changelog:lint`.
3. Bump `package.json` to the prerelease version, for example `0.11.3-beta.1` or `0.11.3-rc.1`.
4. Run the prerelease gate locally:
   `bun run changelog:prerelease-notes --version <version>`
   `bun run verify:config-example`
   `bun run typecheck`
   `bun run test:fast`
   `bun run test:env`
   `bun run build`
5. Commit the prerelease prep. Do not run `bun run changelog:build`.
6. Tag the commit: `git tag v<version>`.
7. Push commit + tag.

Prerelease tags publish a GitHub prerelease only. They do not update `CHANGELOG.md`, `docs-site/changelog.md`, or the AUR package, and they do not consume `changes/*.md` fragments. The final stable release is still the point where `bun run changelog:build` consumes fragments into `CHANGELOG.md` and regenerates stable release notes.

Notes:

- Versioning policy: SubMiner stays 0-ver. Large or breaking release lines still bump the minor number (`0.x.0`), not `1.0.0`. Example: the next major line after `0.6.5` is `0.7.0`.
- Supported prerelease channels are `beta` and `rc`, with versions like `0.11.3-beta.1` and `0.11.3-rc.1`.
- Pass `--date` explicitly when you want the release stamped with the local cut date; otherwise the generator uses the current ISO date, which can roll over to the next UTC day late at night.
- `changelog:check` now rejects tag/package version mismatches.
- `changelog:prerelease-notes` also rejects tag/package version mismatches and writes `release/prerelease-notes.md` without mutating tracked changelog files.
- `changelog:build` generates `CHANGELOG.md` + `release/release-notes.md` and removes the released `changes/*.md` fragments.
- In the same way, the release workflow now auto-runs `changelog:build` when it detects unreleased `changes/*.md` on a tag-based run, then verifies and publishes.
- Do not tag while `changes/*.md` fragments still exist.
- Prerelease tags intentionally keep `changes/*.md` fragments in place so multiple prereleases can reuse the same cumulative pending notes until the final stable cut.
- If you need to repair a published release body (for example, a prior version’s section was omitted), regenerate notes from `CHANGELOG.md` and re-edit the release with `gh release edit --notes-file`.
- Prerelease tags are handled by `.github/workflows/prerelease.yml`, which always publishes a GitHub prerelease with all current release platforms and never runs the AUR sync job.
- Tagged release workflow now also attempts to update `subminer-bin` on the AUR after GitHub Release publication.
- AUR publish is best-effort: the workflow retries transient SSH clone/push failures, then warns and leaves the GitHub Release green if AUR still fails. Follow up with a manual `git push aur master` from the AUR checkout when needed.
- Required GitHub Actions secret: `AUR_SSH_PRIVATE_KEY`. Add the matching public key to your AUR account before relying on the automation.
