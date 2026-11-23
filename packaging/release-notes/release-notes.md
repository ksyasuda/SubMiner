## Highlights
### Added
- **Outro Skipping**:
  - The AniSkip key now skips endings too, with its own prompt and chapter markers.
  - Intro and outro detection reads the file's own chapters first, so it also works with Jellyfin and other streams.
  - Skip times now match your TV or Blu-ray cut more closely.
  - AniSkip rate limits and network errors no longer leave an episode marked as "no intro" until restart.
- **Anime Franchise Cards**:
  - The stats Library now shows each anime once, with its related seasons, movies, and OVAs grouped on one card.
  - Open a card to see each part in release order with its own stats.
  - A title's page has options to move it to another anime, show it on its own, or restore AniList's grouping.
- **FFmpeg Download in Setup**: First-run setup now shows whether FFmpeg and FFprobe are available. If they're missing, **Download FFmpeg** installs a verified build that SubMiner and the `subminer` command use automatically.

### Fixed
- **AniList Season Matching**: Later seasons now update the correct AniList entry when the series includes split seasons, specials, or movies, as with Mushoku Tensei S3 and Re:ZERO S4. Episodes in the second half of a split season update the right entry and episode number.
- **Jimaku**: The Jimaku modal fills in the series name, season, and episode again for Jellyfin episodes whose titles contain a "/".
- **mpv Keys in the Overlay**: Keys bound in `input.conf`, such as arrow keys for subtitle seeking, are no longer sent to a Yomitan popup that has already closed. They are also no longer dropped when the overlay regains focus.
- **Windows Overlay**:
  - The overlay attaches to mpv more reliably and no longer relies on PowerShell lookups.
  - The mpv IPC socket can now be set in `mpv.conf`.
  - The overlay can no longer stay attached to an unrelated window after mpv restarts.
  - A stalled overlay now logs warnings by default, without debug logging.

### Docs
- **Windows Overlay Troubleshooting**: Added a guide to collecting and reading Windows overlay attachment and readiness warnings.

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
