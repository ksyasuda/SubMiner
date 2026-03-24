## Highlights
### Changed
- Release: Reduced packaged release size by excluding duplicate `extraResources` payload and pruning docs, tests, sourcemaps, and other source-only files from Electron bundles.

### Fixed
- Overlay: Restored controller navigation and lookup/mining controls while the subtitle sidebar is open, while keeping true modal dialogs blocking controller actions.
- Tokenizer: Fixed subtitle annotation clearing so explanatory contrast endings like `んですけど` are excluded consistently across the shared tokenizer filter and annotation stage.

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
