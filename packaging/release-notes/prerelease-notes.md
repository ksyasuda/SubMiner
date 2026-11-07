> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

<!-- prerelease-version: 0.22.1-beta.1 -->

## Highlights
### Fixed
- **Windows Overlay**:
  - The overlay now finds mpv more reliably, without relying on PowerShell lookups.
  - You can now configure the IPC socket in `mpv.conf`.
  - The overlay no longer stays attached to an unrelated window after mpv's process ID is reused.
  - Stalled overlays now log diagnostic warnings without debug logging enabled.

### Docs
- **Windows Overlay Troubleshooting**: Added a guide to collecting and interpreting overlay attachment and readiness warnings without debug logging.

## What's Changed

- fix(overlay): identify Windows mpv windows through IPC PID by @ksyasuda in #330

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
