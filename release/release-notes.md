## Highlights
### Breaking Changes
- **Notification Type `both`**: This setting now routes to overlay + system notifications instead of mpv OSD + system.
  - Set `notificationType` to `osd-system` in `config.jsonc` to keep the previous OSD + system behavior.
  - `osd` and `osd-system` remain valid config-file values but no longer appear as options in the Settings UI.

### Added
- **Overlay Notifications**: A new in-app notification stack replaces bare OSD text for most alerts, using Catppuccin Macchiato styling with 3-second auto-dismiss.
  - Position via `notifications.overlayPosition` (top-left, top-center, or top-right; default top-right). Startup, mining, sync, and error alerts queue for the overlay instead of falling back to raw OSD.
  - Mined-card notifications include card thumbnails and an **Open in Anki** button; update-available notifications include a one-click **Update** button.

- **Notification History Panel**: A slide-in panel logging every notification from the current session, toggled with `Ctrl/Cmd+N` (configurable via `shortcuts.toggleNotificationHistory`).
  - Works whether the overlay or mpv has focus; slides in from the same edge as the notification stack.
  - Entries retain thumbnails and action buttons (Open in Anki, etc.) and can be removed individually or cleared all at once.

- **Stats Search**: A new Search tab for real-time subtitle sentence search across your library.
  - Matches by headword with media context; mine directly to sentence cards, word cards, or audio cards.
  - Sentence cards are queued before slow media generation finishes, so the card lands in Anki quickly with audio filled in later.

### Changed
- **AniSkip**: Intro detection now runs in the SubMiner app rather than the mpv plugin.
  - Covers all files in the mpv session including playlist advances; the plugin no longer makes any network calls.
  - `mpv.aniskipEnabled` and `mpv.aniskipButtonKey` hot-reload without restarting playback. Requires SubMiner to be connected to mpv — plugin-only sessions no longer fetch skip windows.

- **Library**: Local and Jellyfin entries are now split by season using folder structure first, filename parsing as fallback.
  - Existing combined-series stats rows are automatically migrated to season-specific entries on startup.
  - Anime detail and cover art refresh immediately after manually changing an AniList entry.

- **Stats — Vocabulary Review**: Hide Known/Hide Kana filters are remembered across sessions; Related Seen Words now matches on shared readings or kanji; duplicate-collapsed exclusions cover all token variants.

- **Stats — Trends**: Reorganized into Activity, Cumulative Totals, Efficiency, Patterns, and Library sections; disambiguated per-period vs. cumulative charts; added Words/Min and Cards/Hour efficiency charts.

- **Stats — Library Browsing**: Remembers card size between sessions; retries stored cover art preserving PNG/WebP MIME types; honors custom AnkiConnect URLs for Browse; session deletes show progress and refresh faster.

- **Stats Mining**: Several reliability improvements when mining from Search and vocabulary examples.
  - Empty `ankiConnect.deck` falls back to Yomitan's configured mining deck; secondary subtitle auto-selection prefers regular English tracks over Signs/Songs tracks.
  - Invalid stored timings and out-of-order subtitle pairs are skipped before FFmpeg runs; partial media failures are shown inline rather than silently dropped.

### Fixed
- **AniList**: Entries are now marked completed when a post-watch sync reaches the final known episode of the season.
- **AniSkip**: Fixed intro markers disappearing after same-media mpv reloads; fixed detection for intros starting at 0 seconds and common release-group filenames.
- **Jellyfin**: Session restarts after setup login so the websocket reconnects with fresh credentials; session stops on logout.
- **Anki — Sentence Cards**: Generated audio is written only to the configured sentence audio field and no longer also fills the expression audio field.
- **Stats Mining**: Word audio uses configured Yomitan sources; English subtitle text is no longer written to word cards; sentence clips correctly update the SentenceAudio field.
- **Overlay Startup**: Subtitle bars are hoverable and clickable as soon as the first subtitle line appears; Linux overlay input is primed from the first measured surface so first-line subtitles and startup notifications are immediately clickable; an OSD spinner now shows from mpv connect through to content-ready.
- **Startup Autoplay**: SubMiner now releases playback after tokenization and overlay content are ready even when playback begins before the first subtitle line appears.

<details>
<summary>Internal changes</summary>

### Internal
- Release notes now credit contributors with a What's Changed list and a New Contributors section, resolved from changelog fragments via git and the GitHub API.
- Updated `make deps` so a fresh source checkout initializes submodules before installing root, stats, and texthooker-ui dependencies.
- Changed PR changelog guidance to preserve multiple fragments for genuinely separate outcomes and direct contributors to consolidate same-PR churn before merging.

</details>

## What's Changed

- feat(notifications): add overlay notifications with position config by @ksyasuda in #110
- feat(stats): speed up session maintenance and improve stats UI by @ksyasuda in #111
- [codex] Restart Jellyfin remote session after setup login by @bee-san in #112
- docs(changelog): require reconciled fragments, not just new ones by @ksyasuda in #113
- feat(release): add contributor attribution to release notes by @ksyasuda in #114
- fix(anilist): mark entry completed when final episode is reached by @ksyasuda in #115
- feat(aniskip): move intro detection from mpv plugin to app runtime by @ksyasuda in #117
- fix(anki): write sentence card audio only to sentence audio field by @ksyasuda in #118

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
