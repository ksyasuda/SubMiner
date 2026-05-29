<!-- read_when: cutting a tagged release or debugging release prep -->

# Releasing

## Prerequisites

- `claude` (Claude Code CLI) installed, on `PATH`, and authenticated.
  `changelog:build` and `changelog:prerelease-notes` invoke
  `claude -p --model sonnet` to merge and rewrite `changes/*.md` fragments into
  a polished, user-facing release body. Either OAuth login (`claude /login`) or
  `ANTHROPIC_API_KEY` works. Install from <https://claude.com/claude-code> if
  you don't already have it.

## Stable Release

1. Confirm `main` is green: `gh run list --workflow CI --limit 5`.
2. Confirm release-facing docs are current: `README.md`, `changes/*.md`, and any touched `docs-site/` pages/config examples.
3. Run `bun run changelog:lint`.
4. Bump `package.json` to the release version.
5. Build release metadata before tagging (this calls `claude -p` locally):
   `bun run changelog:build --version <version> --date <yyyy-mm-dd>`
   - The polished `CHANGELOG.md` and `release/release-notes.md` are committed
     before tagging. Release CI no longer auto-builds the changelog; it fails
     fast if `changes/*.md` fragments are still present on a tag-based run.
6. Review `CHANGELOG.md` and `release/release-notes.md`. Edit by hand if Claude
   missed something — the committed Markdown is what ships.
7. Run release gate locally:
   `bun run changelog:check --version <version>`
   `bun run verify:config-example`
   `bun run typecheck`
   `bun run test:fast`
   `bun run test:env`
   `bun run build`
   When validating auto-update metadata, also run the relevant platform package
   build and confirm `release/` contains the generated updater metadata
   (`latest*.yml`) and blockmaps (`*.blockmap`).
8. If `docs-site/` changed, also run:
   `bun run docs:test`
   `bun run docs:build`
   `bun run docs:build:versioned`
9. Commit release prep.
10. Tag the commit: `git tag v<version>`.
11. Push commit + tag.

## Prerelease

1. Confirm release-facing docs and pending `changes/*.md` fragments are current.
2. Run `bun run changelog:lint`.
3. Bump `package.json` to the prerelease version, for example `0.11.3-beta.1` or `0.11.3-rc.1`.
4. Run the prerelease gate locally (this calls `claude -p` locally):
   `bun run changelog:prerelease-notes --version <version>`
   `bun run verify:config-example`
   `bun run typecheck`
   `bun run test:fast`
   `bun run test:env`
   `bun run build`
   When validating packaged updater output, confirm the platform build writes
   `latest*.yml` and `*.blockmap` files under `release/`.
5. Commit the prerelease prep (package.json version bump + the generated
   `release/prerelease-notes.md`). CI does not regenerate notes — it uses the
   committed file — so review it before committing. If you add more
   `changes/*.md` fragments for a later beta/RC, rerun
   `bun run changelog:prerelease-notes --version <version>`; the generator uses
   the existing prerelease notes as the baseline and asks Claude to merge only
   the new fragment material. Do not run `bun run changelog:build`.
6. Tag the commit: `git tag v<version>`.
7. Push commit + tag.

Prerelease tags publish a GitHub prerelease only. They do not update `CHANGELOG.md`, `docs-site/changelog.md`, or the AUR package, and they do not consume `changes/*.md` fragments. The final stable release is still the point where `bun run changelog:build` consumes fragments into `CHANGELOG.md` and regenerates stable release notes.
Prerelease tags also do not update `https://docs.subminer.moe/`.

Notes:

- Versioning policy: SubMiner stays 0-ver. Large or breaking release lines still bump the minor number (`0.x.0`), not `1.0.0`. Example: the next major line after `0.6.5` is `0.7.0`.
- Supported prerelease channels are `beta` and `rc`, with versions like `0.11.3-beta.1` and `0.11.3-rc.1`.
- Pass `--date` explicitly when you want the release stamped with the local cut date; otherwise the generator uses the current ISO date, which can roll over to the next UTC day late at night.
- `changelog:check` now rejects tag/package version mismatches.
- `changelog:prerelease-notes` also rejects tag/package version mismatches and writes `release/prerelease-notes.md` without mutating tracked changelog files. When that file already exists, the generator includes it in the Claude prompt so later beta/RC notes reuse the reviewed text instead of starting over.
- `changelog:build` generates `CHANGELOG.md` + `release/release-notes.md` (both polished by `claude -p`) and removes the released `changes/*.md` fragments. The CHANGELOG keeps internal notes inside a `<details><summary>Internal changes</summary>` collapse; the release notes drop them entirely.
- The release workflow no longer auto-runs `changelog:build`. If pending `changes/*.md` fragments are present on a tag-based run, CI exits with a clear `::error::` pointing at the local fix. Run `bun run changelog:build --version <version>` locally, commit the polished output, then tag.
- Do not tag while `changes/*.md` fragments still exist.
- Prerelease tags intentionally keep `changes/*.md` fragments in place so multiple prereleases can reuse the same cumulative pending notes until the final stable cut. `make clean` preserves `release/prerelease-notes.md` while deleting generated build artifacts.
- If you need to repair a published release body (for example, a prior version’s section was omitted), regenerate notes from `CHANGELOG.md` and re-edit the release with `gh release edit --notes-file`.
- Prerelease tags are handled by `.github/workflows/prerelease.yml`, which always publishes a GitHub prerelease with all current release platforms and never runs the AUR sync job.
- Tagged release workflow now also attempts to update `subminer-bin` on the AUR after GitHub Release publication.
- Stable release tags update `https://docs.subminer.moe/` and `https://docs.subminer.moe/v/<version>/` through `.github/workflows/docs-pages.yml`; `/main/` continues to show development docs from `main`.
- Keep Cloudflare Pages Git auto-deploy disabled for `docs.subminer.moe`. Production docs are direct-uploaded by Wrangler from GitHub Actions with `--branch main`.
- AUR publish is best-effort: the workflow retries transient SSH clone/push failures, then warns and leaves the GitHub Release green if AUR still fails. Follow up with a manual `git push aur master` from the AUR checkout when needed.
- Required GitHub Actions secret: `AUR_SSH_PRIVATE_KEY`. Add the matching public key to your AUR account before relying on the automation.
- Release and prerelease workflows upload updater metadata (`latest*.yml`) and blockmaps (`*.blockmap`) alongside platform artifacts. Do not remove those files while `electron-updater` is enabled.
- macOS tray app updates use the standard `electron-updater`/Squirrel path. Keep `latest-mac.yml`, the macOS `SubMiner-<version>-mac.zip`, and ZIP blockmap published; Squirrel uses the ZIP payload even when the DMG remains the user-facing installer.
- macOS update metadata and full ZIP downloads are routed through `/usr/bin/curl` before Squirrel installation to avoid Electron main-process network crashes on update checks.
- Windows tray app updates use the standard `electron-updater`/NSIS path. Keep `latest.yml`, the Windows NSIS installer, and installer blockmap published; updater HTTP is routed through main-process fetch to avoid Electron main-process network crashes during update checks.
- Build config emits distinct ZIP names: `SubMiner-<version>-mac.zip` for the macOS Squirrel updater payload and `SubMiner-<version>-win.zip` for the Windows portable fallback. The user-facing DMG and Windows installer keep the unqualified `SubMiner-<version>` basename.
- Linux GitHub release metadata and asset downloads also use `/usr/bin/curl` instead of Electron networking for the same reason.
- Local macOS build-output apps outside `/Applications` or `~/Applications` skip native update checks. Manual tray and launcher checks still use GitHub release metadata to report newer releases, but automatic notifications stay quiet when native app installation is unsupported. To validate auto-update end to end, install the signed and notarized app bundle into one of those Applications folders and point it at a published updater feed.
- The first updater-enabled release cannot update older installs automatically. Users need one manual install to get the updater code.
- Stable auto-update checks ignore beta/RC prereleases by default. Set `updates.channel` to `"prerelease"` on a test install when validating beta/RC updater behavior.
