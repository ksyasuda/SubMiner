## Highlights
### Changed
- Overlay: Expanded the `Alt+C` controller modal into an inline config/remap flow with preferred-controller saving and per-action learn mode for buttons, triggers, and stick directions.

### Internal
- Workflow: Hardened the `subminer-scrum-master` skill to explicitly answer whether docs updates and changelog fragments are required before handoff.
- Release: Automate `subminer-bin` AUR package updates from the tagged release workflow.

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
