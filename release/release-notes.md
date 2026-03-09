## Highlights
### Changed
- Release: Publish unsigned Windows `.exe` and `.zip` artifacts directly from release CI instead of routing them through SignPath.
- Release: Added `bun run build:win:unsigned` for explicit local unsigned Windows packaging.

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
