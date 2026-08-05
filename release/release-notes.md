## Highlights
### Changed

- **Subsync Reference & Target Picker**
  - You can now choose both sides of a sync run: which subtitle is the timing reference and which one gets retimed.
  - The video file itself can be used as the reference for local files (audio-based sync), though a subtitle track stays the default.
  - Works for both alass and ffsubsync, and retiming the secondary subtitle track no longer overwrites your primary one.

### Fixed

- **Startup Logging**
  - Background startup now respects your configured log level even when no `--log-level` flag is passed.
- **Streaming Subtitle Tokenization**
  - Jellyfin playback now seeds tokenization straight from the downloaded subtitle file, so episodes no longer fall back to slow, line-by-line tokenizing while waiting on playback events.
  - Subtitle cues are no longer dropped when switching to a subtitle track embedded in the stream.
  - Prefetching now runs through the whole episode instead of stopping once the cache filled, and the cache clears between episodes so slowdowns don't carry over to later titles.
  - The tokenization cache was expanded from 256 to 2,500 lines, leaving more room for repeated lines (like openings and endings) to stay cached across episodes.
- **Subtitle Line Display**
  - Subtitle lines now appear immediately at their cue time even if tokenization hasn't finished, upgrading in place with annotations once ready.
  - A failed tokenization attempt is no longer cached as plain text, so the line gets another chance at full annotations later.

## What's Changed

- feat(subsync): add reference and target subtitle track picker by @ksyasuda in #181
- fix(logging): surface subtitle processing debug/warn logs by @ksyasuda in #182
- fix(streaming): keep subtitle tokenization prefetch warm for full episodes by @ksyasuda in #183
- fix(overlay): show plain subtitle line immediately on tokenization cache miss by @ksyasuda in #184

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
