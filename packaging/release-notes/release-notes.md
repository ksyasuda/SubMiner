## Highlights
### Changed
- **Stats Session View**: The expanded session view in the stats dashboard has been redesigned. It shows stat tiles for pace, lines, known-word rate, lookups, and paused time, plus a reading density chart of known and unknown words with lookup and card-mined markers.

### Fixed
- **Subtitle Delay**: Retiming subtitles in mpv (z/x or sub-step) now carries through SubMiner. Mined card audio, images, and timestamps match the retimed line, and so do the subtitle sidebar, subtitle navigation, and replay.
- **Subtitle Sidebar**: mpv keys and mouse buttons, such as volume keys and double-click fullscreen, work again while the sidebar is open.

### Docs
- **README**: Refreshed the anime screenshots using the current build.

## What's Changed

- fix(overlay): forward mpv input while subtitle sidebar is open by @ksyasuda in #300
- fix(subtitles): follow mpv sub-delay in mining, sidebar, and navigation by @ksyasuda in #302
- feat(stats): redesign expanded session view with density chart by @ksyasuda in #304

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
