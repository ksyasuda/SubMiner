## Highlights
### Added
- **Japanese Subtitle Generation**:
  - Generate Japanese SRT subtitles locally with whisper.cpp. Start it from a new modal (Ctrl+Shift+G), from the generate button in an empty subtitle sidebar, or with `subminer generate-subs`.
  - Generation shows progress, can be cancelled, and loads the finished subtitles into mpv automatically.
  - Pick an official multilingual Whisper model, including quantized variants, with size and accuracy guidance. You can download it in the app or point Settings at a model you already have.
  - `large-v3-turbo` is recommended when CUDA support is detected, and `small` otherwise.
  - whisper-cli, ffmpeg, and ffprobe are found on PATH unless you override them. SubMiner names any missing tools before a download starts.
  - An optional "Focus on spoken dialogue" mode uses a separately downloaded Silero VAD model. It keeps audible sections it is unsure about, so dialogue under music is not dropped, but songs may also be transcribed.
  - Long passages are split near speech starts or quiet pauses to reduce subtitles that appear too early. When an eligible embedded or external subtitle track is loaded in mpv, it guides the split points.
  - Each passage runs in a fresh Whisper process, which prevents repeated-character output.
- **Subtitle Selection Modal**:
  - An optional modal for choosing primary and secondary mpv subtitle tracks.
  - Turn it on in Settings under Behavior, then press g followed by s to open it. Turning it off restores mpv's own subtitle selection binding.
  - Single-key actions take priority over configured key sequence prefixes.
  - Conflicting sequences are disabled with a warning, and the existing y commands stay reserved.
- **Subtitle Sidebar Copy**:
  - Select dialogue across several sidebar rows and copy it without timestamps using Ctrl/Cmd+C or the Copy button.
  - Selecting text does not seek playback and does not require mining a card.
- **Media Timing Review Screenshot Picker**:
  - Choose the still screenshot separately from the audio range, with a live preview and its own time slider.
  - Step through decoded frames one at a time to get the exact frame you want.
  - Works with local video and with seekable remote streams such as Jellyfin.
- **mpv Keybindings in the Overlay**:
  - The overlay now picks up keyboard bindings from mpv defaults, `input.conf`, and loaded scripts when they do not conflict with SubMiner.
  - SubMiner controls and bindings you explicitly disabled take precedence.
  - These bindings apply only to the current session and are not listed in the help menu.
- **Jimaku Live Action Search**: The Jimaku modal has new Anime and Live action tabs, so you can search Jimaku's live action catalogue as well as anime. Use Arrow Left and Arrow Right to switch tabs.
- **TMDB Live-Action Library**:
  - Live-action dramas and movies in the stats Library now get posters, synopses, and titles from TMDB.
  - Release builds include a project key. Setting `tmdb.apiKey` or `tmdb.apiKeyCommand` overrides it, and one of them is required when running from source.
  - Titles that AniList cannot match are looked up on TMDB automatically when the parsed filename exactly matches a Japanese live-action title. For everything else, use the new **Link to TMDB** action.
  - Entries linked to the same TMDB title merge into one card, and the Library kind selector has a new Live Action option.
  - If a replacement download fails during provider reassignment, the previous link and artwork are kept. Merges and sync keep AniList and TMDB identities separate, and the merge dialog explains mixed selections instead of failing.
- **YouTube Library Kind**:
  - YouTube channels are now their own Library media kind. Existing channel entries migrate automatically, and viewing history and manual video assignments are unchanged.
  - New All Titles, Anime, and YouTube Library filters.
  - Channels are excluded from AniList matching, season repair, and duplicate recommendations.
  - Merges and video moves can no longer combine an anime entry with a YouTube channel.

### Changed
- **Launcher Uses Bundled Bun**:
  - Every installed and downloadable launcher now runs on the Bun runtime that ships with SubMiner. A system Bun is no longer needed.
  - Recognized legacy launchers migrate automatically.
  - Windows gets a `subminer.cmd` launcher download.
  - First-run setup is reduced to one optional launcher control. Runtime repair guidance appears only when it is needed.
- **Faster Sync Transfers**:
  - Sync between compatible macOS and Linux machines now uses compressed, incremental rsync transfers.
  - The last snapshot received from each peer is cached, which reduces traffic on later syncs.
  - Machines without a compatible rsync, including Windows, fall back to compressed scp.
  - Older peers still work without the upload cache.
  - Transfers abort after 30 minutes.
- **Stats Server Request Safety**:
  - The stats server now accepts loopback hosts only and rejects requests from browser origins other than its own.
  - Requests that change data must send an `application/json` body. Scripts that POST to the server need to set a JSON content type.
  - The in-app stats overlay now loads from the local server, so it gets the same protection.
  - Dashboards served through a reverse proxy or Tailscale Serve are no longer supported.
- **Smaller Downloads**:
  - Installers and unpacked apps are smaller. Demo media, source maps, TypeScript sources, test fixtures, and unused Koffi binaries are no longer packaged.
  - All windows now share one Japanese UI font.
  - Release builds publish package size reports that compare against the previous release.
- **Bundled Yomitan**: Updated with upstream Yomitan 26.9.8 changes, including historical Japanese kana transformations, Ukrainian language support, and improvements to Anki duplicate searches and audio retrieval.

### Fixed
- **Jellyfin 12 Compatibility**:
  - Playback, subtitle, artwork, and remote-control requests now authenticate with the `ApiKey` query parameter, so the integration works on Jellyfin 12, where legacy authorization is off by default.
  - "Play on SubMiner" stays available. The cast connection now answers keep-alive requests and reconnects when the server stops responding, instead of silently dying after about a minute.
  - The Jellyfin "now playing" bar clears when you close or finish a cast video instead of running on to the end of the episode.
  - Cards mined during Jellyfin playback get the episode title in the misc info field again instead of "Unknown media".
- **Jellyfin Privacy and Playback**:
  - Jellyfin streams no longer leak titles taken from the stream URL, or stream URLs that contain credentials, into metadata lookups, Anki source fields, Discord presence, stats, or AniList retries.
  - Previously cached metadata that contained credentials is cleaned up. Watch history and library assignments are not touched.
  - Jellyfin playback and casting now use your configured mpv executable, so they work when mpv is installed outside PATH. Portable plugins next to that executable are detected.
- **Anki Mining**:
  - New `ankiConnect.fields.wordAudio` setting reads word audio separately from the sentence audio field. This fixes animated images that started moving immediately when `fields.audio` pointed to `SentenceAudio`. Existing animated images need to be regenerated to pick up the fix.
  - Sentence furigana on word cards stays in sync with the full stats-search context and with expanded timing review selections. Stale readings are cleared if regeneration fails.
  - Closing the overlay while media timing review is still loading now cancels the review, resumes playback if the review paused it, and cleans up the hidden preview player.
  - `ankiConnect.media.maxMediaDuration: 0` now means unlimited when mining from the stats dashboard, matching overlay mining.
  - Invalid AnkiConnect, Kiku, and Senren settings are now rejected with a warning and fall back to defaults.
- **Stats Server Stability**:
  - A port conflict is now reported in a status notification instead of crashing SubMiner.
  - Simultaneous startup requests share one server start. Stopping the background server no longer disconnects dashboards open in the foreground.
  - Shutdown waits only a limited time for active requests to finish.
  - Malformed resource IDs, and ID lists with any invalid entries, are rejected before Library changes or cover backfills run.
- **Subtitle Sidebar**:
  - Clicking a cue no longer leaves the row focused, and Space no longer seeks back to a focused cue. Enter still seeks to the focused cue, and Space keeps its configured playback action.
  - The sidebar stays near the current playback position during gaps when the subtitle file has a cue that starts at zero.
- **Settings Save Feedback**:
  - Settings marked LIVE no longer show false restart warnings, including for notifications and subtitle generation.
  - When a save mixes live and restart-only changes, the live changes apply right away and only the changed sections that need a restart are listed.
- **Overlay Windows**:
  - On Hyprland, recovery dialogs stay above SubMiner windows so overlay placement updates no longer cover their Wait and Close buttons.
  - On Linux, a delayed close callback during teardown can no longer reopen the overlay.
- **First Launch on macOS**: SubMiner no longer exits on first launch when the config directory does not exist yet.

### Docs
- **Launcher**: Documented the launcher install that uses the bundled runtime, migration from legacy launchers, package-managed updates, and the bundled Bun runtime's MIT and LGPL notices. The AUR package installs these notices under `/usr/share/licenses/subminer-bin`, and they are also included in `subminer-assets.tar.gz`.
- **Subtitle Generation**: Documented model choice, VAD behavior, splitting guided by a reference track, fallback behavior, and known limits.
- **Subtitle Selection**: Documented the subtitle selector setting, its shortcut override, and the primary and secondary track controls.
- **Settings**: Clarified save feedback for live settings, warnings for saves that mix live and restart-only changes, and how subtitle generation settings reload.
- **Jellyfin**:
  - Clarified that Windows mpv playback and Jellyfin casting can use a configured executable path instead of PATH.
  - Documented how Jellyfin media titles and stats identities keep stream credentials out of metadata.
- **Stats Library**:
  - Documented TMDB linking, provider reassignment, merge compatibility, and caching of the credential command's output.
  - Documented YouTube channel filtering and video statistics in the Library.
- **Mining**:
  - Documented choosing the screenshot separately in media timing review.
  - Documented the separate word audio field mapping, including that existing animated images need to be regenerated.
- **Sync**: Documented compressed transfers, where the incremental sync cache is stored, and compatibility with older peers.

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
