> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

<!-- prerelease-version: 0.22.1-beta.2; since: v0.22.1-beta.1 -->

## Changes since v0.22.1-beta.1

- Later seasons now land on the right AniList entry even when the sequel chain includes split-season halves, specials, or movies (for example Mushoku Tensei S3, Re:ZERO S4, and Attack on Titan S4).
  - Entries whose title names the season ("Season 3", "III", "第3期") are now preferred over counting sequels.
  - Episodes in the second half of a split season now update that half's AniList entry at the correct episode number.
  - Character dictionary AniList matches cached by older versions are re-resolved once.
- Added outro skipping: the AniSkip key now also skips endings, with its own prompt and chapter markers.
- Intro and outro detection now reads the file's own chapters first and only asks AniSkip for what is missing, so it also works for Jellyfin and other streams.
- AniSkip lookups now send the episode length, so skip times match your TV or Blu-ray cut.
- AniSkip rate limits, server errors, and network failures are no longer remembered as "no intro" until restart; the lookup is retried on the next load.
- The Jimaku modal fills in the series name again for Jellyfin episodes whose name contains a "/", and the AniList fallback guess is fixed for the same case.
- The stats Library now shows each anime once, with its linked seasons, movies, and OVAs grouped into one card that lists the parts in release order.
  - A title's page now has options to move it to another anime, show it on its own, or restore AniList's grouping, and `subminer sync` keeps the most recent choice.
- mpv keys from `input.conf` are no longer sent to a Yomitan popup that has already closed.
- mpv keys are no longer dropped while the overlay reloads mpv's bindings after it regains focus.
- With `--log-level debug`, the log now says why the overlay kept an mpv-bound key instead of forwarding it.
- First-run setup now shows whether FFmpeg and FFprobe are usable, and a **Download FFmpeg** button saves a verified build that SubMiner and the `subminer` command use when no system FFmpeg is found.

## Highlights
### Added
- **Intro and Outro Skipping**:
  - The AniSkip key now skips endings too, with its own prompt and chapter markers.
  - Intros and outros are detected from the file's own chapters first, so skipping also works for Jellyfin and other streams.
  - AniSkip times now match your TV or Blu-ray cut.
- **Library Franchise Cards**: Seasons, movies, and OVAs of the same anime now share one library card, with each part's own stats listed inside. To fix the grouping, use **Move to Another Anime…**, **Show On Its Own**, or **Put Back Where AniList Has It** on a title's page.
- **FFmpeg Download in Setup**: First-run setup now shows whether FFmpeg and FFprobe are usable. If they are missing, click **Download FFmpeg** to install a verified build that SubMiner and the `subminer` command can use.

### Fixed
- **AniList Seasons**: Later seasons now match the right AniList entry when the series has split-season halves, specials, or movies. Episodes in the second half of a split season now update the right entry.
- **AniSkip**: Rate limits, server errors, and network failures are no longer remembered as "no intro" until restart. SubMiner retries them on the next load.
- **Jimaku**: The modal fills in the series name again for Jellyfin episodes whose name contains a "/".
- **Windows Overlay**:
  - The overlay now finds mpv more reliably, without relying on PowerShell lookups.
  - You can now configure the IPC socket in `mpv.conf`.
  - The overlay no longer stays attached to an unrelated window after mpv's process ID is reused.
  - Stalled overlays now log diagnostic warnings without debug logging enabled.
- **mpv Keys in the Overlay**: Keys bound in `input.conf` are no longer sent to a Yomitan popup that has already closed. They also no longer get dropped while the overlay reloads mpv's bindings after it regains focus.

### Docs
- **Windows Overlay Troubleshooting**: New guide to collecting and reading overlay attachment and readiness warnings without debug logging.

## What's Changed

- fix(overlay): identify Windows mpv windows through IPC PID by @ksyasuda in #330
- fix(overlay): keep forwarding mpv keys after a Yomitan popup closes by @ksyasuda in #332
- feat(setup): download FFmpeg from first-run setup when none is usable by @ksyasuda in #333
- feat(aniskip): detect intros and outros from file chapters and AniSkip v2 by @ksyasuda in #334
- fix(anilist): resolve later seasons past split-cour halves and specials by @ksyasuda in #335
- feat(stats): group library seasons, movies and OVAs into franchise cards by @ksyasuda in #336
- fix(jimaku): keep the series name for stream titles containing slashes by @ksyasuda in #337

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz`, the `subminer` launcher, and the Windows `subminer.cmd` launcher
- Bun corresponding source: `bun-v1.3.5-source.tar.gz` and its `.sha256` file

Both launcher downloads use Bun included with the SubMiner app. Download `subminer` on Linux or macOS and `subminer.cmd` on Windows.

The app bundles an unmodified Bun 1.3.5 runtime. Bun is MIT licensed and statically links JavaScriptCore (LGPL 2.0) and TinyCC (LGPL 2.1). License texts and third-party notices ship inside the app under `resources/bun/licenses`, and the source archive above contains the matching Bun, WebKit, and dependency sources for relinking.
