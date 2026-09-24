> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

<!-- prerelease-version: 0.20.0-beta.2; since: v0.20.0-beta.1 -->

## Changes since v0.20.0-beta.1

- Added an optional subtitle selection modal for primary and secondary mpv subtitle tracks. Enable it in Settings > Behavior, then press `g` followed by `s`. Disabling it restores mpv's subtitle selection binding.
  - Single-key shortcut actions now take priority over configured multi-key sequence prefixes, conflicting sequences are disabled with a warning, and existing `y` commands remain reserved.
- Jellyfin casting and playback now honor a configured mpv executable path, allowing playback when mpv is installed outside the system PATH, and portable plugins located beside that executable are now detected.
- Fixed word-card sentence furigana falling out of sync with full stats-search context and expanded timing-review selections; stale furigana is now cleared if regeneration fails.

## Highlights

### Added

- **Japanese Subtitle Generation**:
  - Generate Japanese subtitles locally with whisper.cpp, right from a modal (`Ctrl+Shift+G`), the subtitle sidebar's generation button when no subtitles are loaded, or `subminer generate-subs`, with progress, cancellation, and automatic loading into mpv when it's done.
  - Pick and download an official multilingual Whisper model in-app (including smaller quantized variants), or point Settings at one you already have. SubMiner recommends `large-v3-turbo` when CUDA is available and `small` otherwise, and tells you up front if `whisper-cli`, `ffmpeg`, or `ffprobe` can't be found.
  - An optional "Focus on spoken dialogue" mode uses a Silero VAD model to keep quiet or music-covered dialogue that would otherwise get dropped.
  - Long passages split near natural speech pauses, guided by an existing subtitle track when one is loaded, giving tighter timing and fewer repeated-word glitches.

- **Media Timing Review Frame Picker**:
  - The screenshot used for a mined card can now be chosen independently of the audio clip, with its own live preview, time slider, and frame-by-frame stepping.
  - Works for local video and for seekable remote streams like Jellyfin.

- **Overlay Keybinding Pickup**: The overlay now recognizes your mpv keybindings (from mpv's defaults, `input.conf`, and loaded scripts) as long as they don't conflict with SubMiner's own controls. Picked-up bindings work for the session but won't show up in the help menu.

- **Subtitle Selection Modal**:
  - An optional subtitle selection modal lets you pick primary and secondary mpv subtitle tracks without leaving the overlay.
  - Enable it in Settings under Behavior, then trigger it with `g` followed by `s`; turning it off restores mpv's normal subtitle selection binding.
  - Single-key shortcuts always take priority over multi-key sequences, and any conflicting sequence is disabled with a warning instead of misbehaving.

- **Subtitle Sidebar Selection & Copy**: You can now select dialogue across multiple subtitle sidebar rows and copy it, without timestamps, using Ctrl/Cmd+C or the Copy button, without seeking or mining a card.

- **Jimaku Live Action Search**: The Jimaku modal has separate Anime and Live Action tabs (switch with Arrow Left/Right) so you can search Jimaku's live-action subtitle catalogue directly.

- **Live-Action TMDB Library**:
  - Live-action dramas and movies in the stats Library now get posters, synopses, and titles from TMDB.
  - Titles AniList can't match are looked up on TMDB automatically when the parsed filename matches a title exactly; otherwise use the new **Link to TMDB** action. Entries linked to the same TMDB title merge into one card, and the Library kind selector gained a Live Action option.
  - Release builds already include a TMDB key; if you run from source, set `tmdb.apiKey` (or `tmdb.apiKeyCommand`) yourself.

- **YouTube Library Kind**:
  - YouTube channels are now their own Library media kind, with new All Titles, Anime, and YouTube filters. Existing channel entries migrate automatically with viewing history and manual video assignments intact.
  - Channels stay out of AniList matching, season repair, and duplicate recommendations, and can't be merged or moved into an anime entry.

### Changed

- **Bundled Bun Runtime**: Every SubMiner launcher, installed or downloaded, now runs on the Bun runtime bundled with the app instead of a system-wide Bun install. Recognized legacy launchers migrate automatically, Windows users get a new `subminer.cmd` download, and first-run setup now shows a single optional launcher control with runtime repair guidance only when something actually needs it.

- **Compressed Incremental Sync**: Cross-machine sync between compatible macOS/Linux machines now transfers only what changed, compressed, using a cached snapshot from the last sync to cut traffic further. Machines without a compatible rsync (including Windows) fall back to compressed scp automatically, older peers keep working, and transfers now time out after 30 minutes instead of hanging indefinitely.

- **Smaller Install Size**: Installers and the unpacked app are smaller after dropping demo media, source maps, TypeScript sources, test fixtures, and unused binaries, and sharing one Japanese UI font across windows. Release builds now publish a package-size comparison against the previous release.

- **Stats Server Request Safety**: The stats server, including the in-app stats overlay which now loads through it, only accepts requests from the local machine and requires a JSON content type for anything that changes data. If you were exposing the dashboard through a reverse proxy or Tailscale Serve, that's no longer supported, and any script posting to the stats API needs to send `Content-Type: application/json`.

- **Yomitan Updated**: Bundled Yomitan is updated to upstream 26.9.8, adding historical Japanese kana transformations and Ukrainian language support, plus improvements to Anki duplicate search and audio retrieval.

### Fixed

- **Jellyfin**:
  - Playback, subtitles, artwork, and remote control now authenticate with an `ApiKey` parameter instead of legacy headers, so Jellyfin 12 works correctly even with legacy authorization disabled.
  - "Play on SubMiner" no longer silently drops the connection after about a minute on Jellyfin 12.
  - Casting now honors your configured mpv executable path, so playback works and portable plugins are detected correctly when mpv isn't on PATH.
  - The "now playing" bar clears when you close or finish a cast video instead of running to the end of the episode.
  - Anki cards mined from Jellyfin now get the real episode title in the misc info field instead of "Unknown media".
  - Jellyfin streams no longer leak URL-derived titles or credential-bearing URLs into metadata, Anki fields, Discord presence, stats, or AniList lookups; previously cached data that had credentials in it is cleaned up automatically.

- **Anki & Mining**:
  - Word audio now reads from its own configured field (`ankiConnect.fields.wordAudio`) instead of the sentence-audio field, fixing animated word images that started moving immediately instead of on demand.
  - Setting `ankiConnect.media.maxMediaDuration` to `0` for unlimited duration now also applies when mining from the stats dashboard, matching overlay mining.
  - Closing the overlay while a media timing review is still loading now properly cancels setup, restores playback if the review had paused it, and cleans up the hidden preview player.
  - Word-card sentence furigana now stays in sync with the full stats-search context and expanded timing-review selections, and clears stale readings automatically if regeneration fails.

- **Settings**:
  - AnkiConnect, Kiku, and Senren settings are now validated before use, with a warning and a safe default for anything invalid instead of a bad value reaching runtime.
  - Settings marked as applying live now correctly avoid showing a restart warning, and mixed saves apply the live parts immediately while listing only the sections that actually need a restart.

- **Overlay**:
  - Clicking a subtitle sidebar cue no longer leaves Space bound to seeking back to it; Enter still seeks the focused cue, and Space keeps whatever playback action you've configured.
  - Hyprland recovery dialogs now stay above SubMiner windows instead of being covered by overlay placement updates.
  - Fixed a rare case on Linux where a delayed window-close callback could reopen the overlay after it was torn down.

- **Stats**:
  - Malformed or partly invalid resource IDs are now rejected before they can affect library mutations or cover-art backfills.
  - Stats server port conflicts now surface as a status notification instead of crashing SubMiner, and startup/shutdown are more robust: concurrent startup requests share one attempt, stopping a background instance no longer disconnects an open dashboard, and shutdown no longer waits indefinitely on active requests.

- **Subtitle Sidebar Gap Follow**: The subtitle sidebar now stays near actual playback position during gaps in files where a cue starts at time zero.

- **First Launch on macOS**: Fixed first launch exiting immediately when the SubMiner config directory didn't exist yet.

## What's Changed

- feat(sidebar): add dialogue selection and copying by @ksyasuda in #238
- feat(subtitles): add local Japanese subtitle generation by @ksyasuda in #240
- perf(stats): use compressed incremental snapshot transfers by @ksyasuda in #241
- fix(startup): create config directory before singleton lock by @ksyasuda in #242
- feat(launcher): bundle Bun and use it across all launchers by @ksyasuda in #243
- build(release): reduce package size and report release sizes by @ksyasuda in #244
- fix(overlay): keep Hyprland recovery dialogs above overlays by @ksyasuda in #245
- feat(overlay): discover unclaimed mpv key bindings by @ksyasuda in #246
- fix(sidebar): preserve Space playback after cue seeking by @ksyasuda in #247
- fix(jellyfin): fix jellyfin media metadata by @ksyasuda in #250
- feat(jimaku): add live-action subtitle search by @ksyasuda in #251
- feat(stats): add TMDB metadata for live-action dramas in the Library by @ksyasuda in #252
- feat(stats): separate YouTube channels in the Library by @ksyasuda in #253
- feat(mining): add a screenshot frame picker to media review by @aalhendi in #254
- fix(config): align live save feedback with hot reload policy by @ksyasuda in #255
- fix(anki): separate word audio mapping for animation sync by @ksyasuda in #256
- fix(config): validate AnkiConnect and field grouping settings by @ksyasuda in #257
- fix(anki): honor unlimited duration in stats mining by @ksyasuda in #258
- fix(stats): reject malformed resource IDs before mutations by @ksyasuda in #259
- fix(stats): harden server lifecycle and verify compiled runtime by @ksyasuda in #261
- fix(overlay): cancel pending window transitions and timing reviews by @ksyasuda in #262
- fix(stats): restrict local requests and serve the dashboard over HTTP by @ksyasuda in #263
- fix(jellyfin): support modern authentication by @ksyasuda in #264
- feat(overlay): add optional subtitle selection modal by @ksyasuda in #265
- fix(jellyfin): respect Windows mpv configuration when casting by @aalhendi in #267
- fix(anki): regenerate sentence furigana from the final sentence by @ksyasuda in #268

## New Contributors

- @aalhendi made their first contribution in #254

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
