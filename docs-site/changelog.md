# Changelog

## v0.19.5 (2026-08-30)

**Fixed**

- **Anki Card Update Progress**: The card-update spinner now stays visible until audio and image updates finish, instead of disappearing early.
- **Anki Word-Card Fields**: Word-card enrichment now writes sentence text and audio to the fields configured in AnkiConnect, while the dedicated sentence-card and audio-card actions keep their existing compatible field names.
- **Overlapping Subtitles**:
  - Subtitle lines that start while another line is still on screen now appear alongside it, instead of staying hidden until a track switch or seek.
  - Subtitles shown at the same time now stack by their authored screen position, with top signs and song lines above bottom dialogue.
  - Half-size ASS furigana is no longer shown as if it were a dialogue line.
- **YouTube Auto Captions**:
  - Auto-generated captions now follow their intended timing and two-row roll-up layout.
  - Long speech is paged instead of covering the video with a wall of text.
  - Explicitly timed sound cues like `[音楽]` no longer cover later dialogue.

## v0.19.4 (2026-08-25)

**Added**
- **Library Merge & Move**: Duplicate library cards for the same show can now be combined. Select cards in the library grid and use "Merge Selected" to pick which entry to keep and move every episode onto it, preserving sessions, mined cards, and watch time. Episodes can also be reassigned individually via the "→" button, useful when a file lands under a stray title; manual assignments survive later filename parsing, Jellyfin refreshes, and season repair. Exact AniList title matches with compatible seasons now merge automatically, while fuzzy matches surface as dismissible "Possible duplicate" reviews instead of merging silently.
- **Duplicate Line Cleanup Tool**: The Vocabulary tab's new "Duplicates" button scans a chosen time window (7 days through all time) for old karaoke/typeset duplicate-line bursts, shows what it found, and collapses each run to one line once confirmed; `subminer stats cleanup --duplicate-lines` does the same from the terminal, with `--dry-run` and `--lookback-days <n>` options. Watch time and lines-seen totals are left unchanged.

**Changed**
- **Prerelease Release Notes**: Prerelease notes now open with a "Changes since" section listing only what changed versus the previous beta/RC of the same version, above the cumulative highlights, and CI rejects prerelease tags whose committed notes were generated for a different beta/RC.

**Fixed**
- **Subtitle & Karaoke Duplication**:
  - Karaoke and animated signs are reconstructed once from their authored text and shown only while actually sung, with original word spacing preserved, instead of flooding the overlay, subtitle sidebar, immersion history, mining, or stats with glyph fragments, per-frame color phases, and repeated animation events.
  - Decorative layers (highlight sweeps, glow/shadow copies, symbol-font decoration, particle swarms, hidden or zero-scaled text) stay out of published text, while ordinary repeated dialogue, positioned signs, wrapped lyric rows, and multi-row CC-style blocks still display correctly.
  - Embedded subtitle tracks on network-mounted (SMB/NFS) media are extracted and parsed again instead of falling back to live-text-only, restoring karaoke reconstruction, sidebar cues, and mining for releases that only ship subtitles inside the container.
  - Secondary subtitles go through the same deduplication pipeline as primary subtitles and no longer clip display after about four lines.
  - Event-heavy karaoke files that previously stalled subtitle loading for several seconds now parse in well under a second.
- **Character Dictionary Reliability**:
  - Generation, merged rebuilds, and imports no longer freeze the app on large dictionaries; snapshot I/O, archive building, and image/name lookup caches moved off the UI's critical path.
  - Dictionaries are reused instead of regenerated when MeCab finds no name splits.
  - Cached portraits restore correctly after the portrait index finishes loading post-tokenization.
  - Desktop progress notifications on Linux AppImage installs update in place instead of flickering, fixing a bug where the AppImage's bundled libraries broke the system notification helper.
- **Overlay Startup & Modals**:
  - The macOS window-tracking helper targets macOS 12.0+ instead of requiring the build machine's exact macOS version, fixing crashes on older systems like Ventura that left the overlay stuck on "Overlay loading".
  - mpv IPC connection attempts time out and retry, showing an actionable error if content still isn't ready after 30 seconds.
  - Dedicated overlay modals are prewarmed on macOS and Windows so shortcuts open them promptly.
  - On macOS, reused modals and the stats window open above fullscreen mpv on its current Space instead of jumping to another desktop.
- **Wayland File Drop**: Fixed native Wayland drag-and-drop from file managers such as Thunar, so subtitle and video files dropped on the visible overlay are resolved and forwarded to mpv.
- **Windows Mouse Lag**: Fixed system-wide mouse lag on Windows while SubMiner is running, caused by the overlay's global mouse hook for click-through forwarding and by the mpv window tracker blocking the app on repeated PowerShell lookups.
- **Sentence Mining Audio & Clips**: Sentence-audio generation no longer times out on slow network-mounted media with many subtitle/font streams (bounded FFmpeg probing, two-minute extraction budget, clearer error reporting), and mined audio/animated AVIF clips now capture the subtitle line that was actually mined by snapshotting the clip range at lookup time instead of reading live mpv state later.
- **Stats Performance & Reliability**: Immersion stats storage now sets its SQLite busy timeout before WAL setup, avoiding transient lock errors under concurrent writes. Deletes in the stats dashboard no longer freeze the UI, run proportional to what's deleted instead of rebuilding full lifetime summaries, retry safely if the delete worker crashes, and no longer rescan the whole library when deleting very common words; a new index also makes large session deletes drop from minutes to milliseconds. Library merges, video moves, and AniList reassignments got the same lifetime-summary fix.
- **Vocabulary Stats Accuracy**: Vocabulary totals and charts now count all tracked vocabulary instead of only the first page, new-word history uses corrected daily rollups (fixing legacy timestamp and time-zone issues), summary cards refresh automatically after edits to the exclusion list, and rapid exclusion edits no longer race each other.
- **Rofi MKV Thumbnails**: Fixed missing MKV thumbnails in the Linux rofi picker when system thumbnailer registrations only advertise legacy Matroska MIME aliases.

<details>
<summary>Internal changes</summary>

**Internal**
- Docs Site Indexing: Excluded the `/main/` and `/v/<version>/` docs trees from search indexing (self-referential canonical, `noindex,follow`, matching `X-Robots-Tag`) so crawlers focus on current docs instead of ~30 archived copies of every page, and restored `<lastmod>` dates in the docs sitemap that were silently dropped by production builds.

</details>

## v0.19.3 (2026-08-13)

**Added**
- Changelog Modal: Adds an in-app changelog you can open from the tray ("View Changelog") or the "What's New" button on the update notification, so the notification stays reachable while you read. It shows the newest published release notes (falling back to the bundled changelog if that fetch fails), folds older versions while keeping the current one expanded, and supports keyboard navigation (`J`/`K`/arrows, `Enter`, `R`, `Esc`).

**Changed**
- Subtitle Tokenization Performance: Reworks subtitle dictionary lookups to cut per-line work roughly in half, cache repeated lookups across lines, and stop tokenization from competing with on-screen subtitle prefetching. Also fixes several accuracy issues along the way: dropped readings on trailing kana, character names being skipped after a dictionary sync, annotations not refreshing after mining a card, and halfwidth katakana character names losing their reading or being swallowed by other words.

**Fixed**
- Character Dictionary Large Imports: Large character dictionaries (e.g. One Piece) no longer fail to install from a fixed timeout budget; the import now scales its time budget to dictionary size and reports detailed progress (page/character counts, image download progress, elapsed time) instead of one static message.
- Stats Delete Responsiveness: Deleting sessions, episodes, or library entries no longer freezes the stats page or an active video player; deletes are now batched into a single transaction.
- Styled Subtitle Cue Parsing: Heavily typeset subtitles (karaoke, signs) no longer flood the subtitle sidebar with garbage; vector drawing commands are no longer shown as text, and duplicate/animation-burst cues now collapse into one.
- X11 mpv Renderer: Fixes an mpv crash on the first fullscreen toggle for X11/XWayland users with `gpu-next` shaders (e.g. ArtCNN), which was caused by X11 mode forcing the legacy OpenGL renderer.
- X11 Overlay Display Scaling: Fixes the overlay appearing oversized and offset from mpv on X11/XWayland under fractional or mixed-monitor display scaling.

<details>
<summary>Internal changes</summary>

**Internal**
- Subtitle text is now decoded from ASS exactly once at ingest, so the renderer, timing tracker, and tokenizer all share one decoded value instead of each re-deriving it.
- Added per-stage debug timings (`scanMs`, `mecabMs`, `frequencyMs`, `annotateMs`) to the subtitle tokenization pipeline log.

</details>

## v0.19.2 (2026-08-04)

**Changed**
- Subsync: The sync modal now lets you choose both the reference subtitle (correct timing) and the out-of-sync subtitle to retime, for both alass and ffsubsync. alass can also use the loaded video's audio as a reference for local files. Retiming the secondary track now reloads the result into the secondary slot instead of overwriting the primary subtitle.

**Fixed**
- Streaming Subtitle Tokenization: Jellyfin streams now seed subtitle tokenization directly from the downloaded subtitle file instead of relying on an mpv event that could be missed, and prefetching now runs to the end of the file and clears between episodes. The tokenization cache was raised from 256 to 2500 lines, and parsed cues are no longer lost when the active subtitle track briefly can't be resolved (e.g. switching to an embedded track). Together these prevent episodes from falling back to slow, line-by-line tokenization during playback.
- Overlay: Subtitle lines now appear immediately at their cue time even on a tokenization cache miss, upgrading in place once tokens and annotations are ready, instead of waiting on a line still being processed. A failed tokenization is no longer cached as plain text, so repeated lines get another chance at annotations.
- Background Logging: Background startup now respects the configured logging level when no explicit log level is passed.

<details>
<summary>Internal changes</summary>

**Internal**
- Patched three high-severity dependency advisories (`undici`, `brace-expansion`, `fast-uri`).

</details>

## v0.19.1 (2026-08-01)

**Added**
- Word Card Type: Adds a setting (Settings > Mining/Anki > Kiku/Lapis Features > "Word Card Type") to choose which card-type flag SubMiner marks on Kiku/Lapis word cards — `word-and-sentence` (default), `click`, `sentence`, `audio`, or `none`. Click cards (`IsClickCard`) can now be flagged, and setting any card-type flag clears the others so a note can't claim two types at once.

**Fixed**
- Yomitan Popup: Fixes the macOS Yomitan popup going inert after mining a card — clicks outside the popup no longer pass through to mpv, and scrolling over the popup scrolls its definitions instead of seeking playback.
- YouTube Playlist Links: Fixes opening a video from a playlist URL (e.g. a Watch Later link with `list=`/`index=`) timing out while probing subtitles, metadata, or the playback URL.

## v0.19.0 (2026-07-29)

**Added**
- Anki Maturity Highlighting: Known-word subtitle highlights can now be colored by Anki card maturity (new, learning, young, mature), similar to asbplayer. Tier thresholds and colors are configurable, with a runtime toggle and an updated help legend.
- Post-Playback Menu: After a watch-history episode ends, the fzf/rofi launcher returns to that series with options to play the previous or next episode, rewatch, pick another episode, or quit. The pre-playback series menu now offers the previous episode too.
- Delete Library Entries: The stats Library detail view can now delete an entire title in one step (episodes, sessions, subtitle lines, rollups, cover art, and vocabulary counts). Delete progress is now shown app-wide via a progress bar and status toast instead of disappearing when you switch tabs.
- Cross-Machine Sync: Added SSH-based syncing of stats and watch history between machines, available from the tray ("Sync Stats & History") or `subminer sync`, with saved devices, per-host sync direction, background auto-sync, connection testing, manual snapshots, and support for Windows remotes.
- TsukiHime Subtitle Downloads: Added subtitle downloads for the current video via TsukiHime, loading Japanese as the primary track and your configured secondary language directly into mpv.

**Changed**
- Clipboard-Video Shortcut: The "append clipboard video to queue" shortcut is now configurable.

**Fixed**
- AniList Season Resolution: Season 2+ files now resolve to the correct AniList entry instead of silently falling back to season 1 (which mismatched character dictionaries and watch progress). Manual overrides now stay scoped per season, fix both the dictionary and progress tracking together, and also correct per-season cover art.
- Subtitle Annotation Accuracy: Fixed several annotation edge cases, including inconsistent POS exclusions on merged quote-particle tokens, dropped annotations on supplementary-plane kanji, katakana punctuation wrongly treated as noise, and certain kanji vocabulary losing N+1 highlighting eligibility.
- AnkiConnect Proxy Port Conflict: Video startup no longer crashes when another process already holds the configured AnkiConnect proxy port; a notification now explains how to resolve it.
- AppImage Quit Crash: Fixed a "Service Crash" desktop notification appearing after closing a video when running the Linux AppImage.
- Autoplay Pause Timing: Fixed playback resuming a few seconds before subtitle tokenization warmup finished, most noticeable when resuming mid-episode or when a cue starts within the first two seconds.
- Stats Known-Word Count: Fixed stats reporting 0 known words for every session after the known-word cache format changed.
- Stats Library Cover Art: Relinking a title to a different AniList entry now updates its cover in the Library grid immediately instead of leaving a stale, mismatched cover cached.
- Rofi Prompt Spacing: Rofi menu prompts now keep a space before the input field instead of running into the placeholder text.
- Stats Settings & Reliability: Hardened stats settings validation (nested/legacy AnkiConnect config now falls back safely instead of breaking) and stats routes against malformed requests and other edge cases.
- Stats Delete Performance: Deleting sessions, episodes, and library entries is now dramatically faster and no longer stalls playback (e.g. a 12-episode title dropped from about a minute to under a second on a large library); the Vocabulary tab also loads much faster. The first launch after upgrading runs a one-time database migration.

<details>
<summary>Internal changes</summary>

**Internal**
- Added a golden-file regression test corpus for the tokenizer/annotation pipeline, plus scripts to record new fixtures and diff against stock Yomitan.
- Consolidated renderer modal state handling into a descriptor registry.
- Consolidated CI quality checks (PR, stable, and prerelease) into one reusable workflow with mpv plugin tests and dependency audits.
- Removed the unused stats IPC transport and unified stats dashboard HTTP types with the backend contract.
- Added a script to verify known-word highlight tiers against live Anki data outside of playback.

</details>

## Previous Versions

<details>
<summary>v0.18.x</summary>

<h2>v0.18.0 (2026-07-10)</h2>

**Added**
- Sentence Audio Normalization: Generated sentence audio is now normalized to -23 LUFS by default, and clips mined from playback mirror mpv's software volume curve with a limiter to prevent clipping. Both behaviors are configurable independently.
- Watch History Command: Added `subminer -H` / `--history` to browse watch history, replay or continue episodes, or pick one via fzf or rofi, with cover art shown in the rofi picker.

**Changed**
- Fzf Preview Layout: Moved fzf previews below launcher menus, giving long titles and metadata more room.
- Known-Word Highlighting: Now compares subtitle and Anki-card readings, preventing false matches between homographs and unrelated words that share a reading, while still supporting matching across kana and kanji spellings.
- Annotation Filtering: Standalone suffix tokens (e.g. さん, れる) are now excluded from JLPT/frequency/N+1 highlighting by default, matching how particles and interjections are treated; configurable via the pos2 exclusion setting.
- App Icon: Replaced the app icon with new pixel-art submarine artwork contributed by the community, used across the app icon, tray, notifications, README, docs site, and Stats page.
- Stats Trend Charts: Overhauled with persisted title visibility, per-chart title limits, "top" and "most recent" ranking modes, an option to show or hide empty days, calendar-aligned periods, and value-sorted tooltips.

**Fixed**
- Background Stats Server: `subminer app` background launches now auto-start the stats server when enabled, and skip startup if one is already running.
- Character Name Highlighting: Character dictionaries now split unspaced native names more reliably, and portraits, highlights, and hover lookup survive punctuation, unmatched text, and competing dictionary matches without incorrectly splitting longer words.
- Highlighting Coverage: Frequency/JLPT highlighting and vocabulary stats now include content adverbs (e.g. 確かに, やはり) and kanji nouns MeCab tags as non-independent (e.g. 日, 点, 以外), while still suppressing interjections, pronouns, and grammar fragments; lexicalized kana expressions like かといって keep their annotations.
- Cover Art Fetching: Stats now fetches AniList cover art as soon as a new series starts playing, and backfills missing art for existing series on the next Stats visit.
- Kiku Field Grouping: The manual field-grouping dialog now stays above fullscreen mpv, remains usable across repeated attempts, closes abandoned windows after timeouts, and reports a clear error when the original card can no longer be loaded.
- Karaoke Subtitle Collapse: Secondary subtitles no longer stack dozens of one-syllable lines during openings and endings; repeated events collapse into one line, capped to a strip at the top.
- Unparsed Token Hover: Subtitle text Yomitan can't parse (truncated inflections, elongation runs) remains hoverable, while staying excluded from highlighting, N+1 candidate math, and vocabulary stats.
- YouTube Streaming: Fixed direct YouTube stream extraction that could corrupt signed URLs and cause ffmpeg 403 errors.

<details>
<summary>Internal changes</summary>

**Internal**
- Test lanes moved to `scripts/test-lanes.ts` with per-directory discovery and isolated per-file timeouts; CI now covers previously orphaned stats, scripts, plugin process-retry, and runtime-compat suites, plus a new stats lane in the change-verification workflow.

</details>

</details>

<details>
<summary>v0.17.x</summary>

<h2>v0.17.2 (2026-06-28)</h2>

**Fixed**
- YouTube Background Cache: Fixed Windows YouTube background media cache startup for YouTube URLs opened directly in mpv, including resolved stream URLs when mpv still exposes the original YouTube playlist entry, so queued Anki media updates can append audio and images after the cache finishes.
- YouTube Subtitle Picker: Manual subtitle picker requests now show an immediate configured notification while SubMiner probes tracks and opens the modal. Subtitle download progress is replaced with a transient success notification after tracks load.

<h2>v0.17.1 (2026-06-27)</h2>

**Added**
- YouTube Media Cache Mode: Adds `youtube.mediaCache.mode` with `direct` and `background` options. Background mode uses a yt-dlp cache download when direct stream extraction is unreliable — creates a text-only card immediately, queues media updates for mined notes, and fills audio/image fields once the download finishes. Progress is announced via overlay/OSD notifications. Downloads are capped at 720p by default (`youtube.mediaCache.maxHeight`). Switching back to direct mode cancels any in-flight background download.

**Fixed**
- Log Export: Fixed log filenames to use the local date so exports around UTC midnight include the current day's logs rather than stale prior-day files. Expanded export redaction to mask IPs, emails, auth and cookie headers, yt-dlp cookie arguments, URL credentials, token/key/password fields, and signed YouTube media URL parameters.
- YouTube Card Media: Improved media generation reliability by sending safer ffmpeg options for resolved streams and skipping stale stream maps (including cached YouTube files). Hardened background cache downloads with IPv4 and extractor retry flags; failed downloads now notify the user and clear queued media updates instead of leaving them silently pending. Stale background cache files are cleaned on startup and before each new download.

<h2>v0.17.0 (2026-06-15)</h2>

**Changed**

- **Subtitle Delay Keybindings**: Updated default overlay subtitle delay and step bindings to match mpv conventions: `z`, `Z`, and `x` adjust `sub-delay`; `Ctrl+Shift+Left/Right` run native `sub-step` with OSD feedback. The previous SubMiner-only adjacent-cue delay action has been removed.
- **Update Notifications**: New installs now default update notifications to overlay-only instead of overlay + system notifications.

**Fixed**

- **Anki – Highlight Word**: Fixed bolding of the mined word in Kiku sentence and sentence-furigana fields when the source Yomitan sentence did not already contain bold markup.
- **Anki – Lapis/Kiku Word Cards**: Fixed word-and-sentence marker missing from Lapis/Kiku word cards enriched through SubMiner, which could hide sentence context on the card front.
- **Anki – Windows Media Generation**: Fixed known-word cache refreshes when no deck is configured, and fixed audio/image generation after background launches by recreating missing FFmpeg temp directories before clipping.
- **Character Dictionary – Windows**: Fixed the Windows "SubMiner mpv" shortcut so character dictionary auto-sync can fall back to mpv's current video path when app media state is not yet ready.
- **Notifications**: Restored the SubMiner app icon on system notifications that do not supply a custom notification image.
- **Overlay – macOS Yomitan Popup**: Fixed Yomitan popup focus after card mining or popup reload; fixed popups staying open when clicking transparent overlay space — click-away now closes the popup and returns click passthrough to mpv without a hide/reappear cycle.
- **Overlay – Linux Auto-Pause Startup**: Fixed the visible overlay on Linux auto-paused startup to stay interactive during the initial measurement gap; startup subtitle cache misses now paint raw text before tokenization finishes, and temporarily empty `sub-text` refreshes are resolved before warm readiness resumes playback.
- **Overlay – Playlist Advance**: Fixed the visible overlay being dismissed when mpv advances to the next playlist item, including when the next episode loads after the warm transition delay.
- **Overlay – Windows Subtitle Bar**: Fixed shaky hover and click interaction on the Windows subtitle bar when a video attaches to an already-running background SubMiner instance.
- **Stats – AniList Linking**: Fixed manual AniList linking from the stats anime page so auto-searches strip the generated "Season N" suffix and query only the anime title.
- **Updates – Linux Support Assets**: Fixed Linux updates to correctly install and refresh the launcher runtime plugin copy and rofi theme alongside AppImage and launcher updates; unrelated SubMiner data directories are now left untouched and plugin copies are staged before replacing the live runtime. Fixed first-launch playback on fresh installs by auto-installing missing managed support assets from the bundled app before mpv starts.

**Docs**

- **Linux Update Flows**: Documented that Linux update flows manage the launcher runtime plugin copy and rofi theme from `subminer-assets.tar.gz`, and that normal launcher playback auto-installs those managed support assets on a fresh install if either is missing.

<details>
<summary>Internal changes</summary>

**Internal**

- **Runtime Modules**: Split main-process runtime wiring into focused modules without changing user-facing behavior; hardened helpers against stale background stats daemon PIDs, stalled subtitle extraction, and dropped async errors.
- **Release CI**: Fixed GitHub release notes to preserve the `What's Changed` and `New Contributors` attribution sections when CI regenerates from the committed changelog; scoped prerelease note reuse to the same base version so a new beta line starts from current fragments.

</details>

</details>

<details>
<summary>v0.16.x</summary>

<h2>v0.16.0 (2026-06-10)</h2>

**Breaking Changes**

- **Notification Mode `both`**: `notificationType: "both"` now routes to overlay + system notifications. Users who previously used `"both"` for mpv OSD + system notifications should set `notificationType` to `"osd-system"` in `config.jsonc`. The `osd` and `osd-system` values remain valid as config-file entries but no longer appear in Settings.

**Added**

- **Overlay Notifications**: New overlay notification stack (Catppuccin Macchiato theme) with configurable screen position (`notifications.overlayPosition`: top-left, top-center, top-right), 3-second transient dismissals, and persistent cards for long-running jobs like character dictionary sync.
- **Notification History Panel**: `Ctrl/Cmd+N` (configurable via `shortcuts.toggleNotificationHistory`) opens a session log of all notifications. Works whether the overlay or mpv has focus; slides in from the notification edge; entries can be individually removed or cleared.
- **Mined-Card Notification Actions**: Mined-card overlay notifications show generated card thumbnails and include an Open in Anki button in both live cards and history entries.
- **Update Notification Action**: Update-available overlay notifications include an Update button to start the app update flow directly from the notification.
- **Stats Search**: New Search tab in Stats for realtime subtitle sentence search with media context, headword matching, and mining actions for source-backed sentence cards or exact-match word/audio cards.

**Changed**

- **AniSkip**: Moved intro detection from the mpv plugin to the SubMiner app. Lookups now cover every file loaded during a session including playlist advances, and `mpv.aniskipEnabled`/`mpv.aniskipButtonKey` hot-reload without restarting playback. The bundled plugin no longer makes network calls. Note: AniSkip now requires the SubMiner app to be connected; plugin-only mpv sessions will not fetch skip windows.
- **Stats Library**: Entries are now split by detected season (season folder first, filename parsing as fallback). Existing combined-series rows are automatically migrated to per-season entries on startup. Cover art and anime details refresh immediately after a manual AniList entry change.
- **Stats Vocabulary**: Remembers Hide Known/Hide Kana filters across sessions, applies Hide Kana filtering cross-title, collapses duplicate token variants in exclusions, and matches Related Seen Words by shared readings or kanji.
- **Stats Trends**: Reorganized into Activity, Cumulative Totals, Efficiency, Patterns, and Library sections; disambiguated per-period vs. cumulative charts; added Words/Min and Cards/Hour efficiency charts.
- **Stats Mining**: Sentence cards are created before slow media generation finishes; stored/requested secondary subtitles are preserved before falling back to sidecar or alass-retimed English subtitles; empty `ankiConnect.deck` falls back to Yomitan's mining deck; partial media failures are surfaced.
- **Stats Browsing**: Remembers library card size; retries stored cover art without extra AniList lookups; preserves PNG/WebP MIME types; honors custom AnkiConnect URLs for Browse; shows progress during session deletes.
- **Startup Notifications**: Tokenization, subtitle annotation, and character dictionary status now route through queued overlay notifications in `overlay`/`both` mode instead of falling back to mpv OSD while the overlay loads.
- **Notification Deduplication**: Cycling subtitle modes updates the active overlay card in place rather than stacking duplicates; repeated progress updates (e.g. subsync) tick in place without flickering.
- **Update Notification Default**: New installs default `notificationType` to `both` so update alerts appear in both overlay and system notifications.

**Fixed**

- **AniList Completion**: Entries are now marked completed when a post-watch update reaches the final known episode of the season.
- **AniSkip Markers**: Fixed intro markers disappearing after same-media mpv reloads; fixed metadata detection for intros that start at 0 seconds and common release-group filenames.
- **Jellyfin Session**: Remote session now restarts after setup login so the websocket reconnects with fresh credentials, and stops cleanly on logout.
- **Sentence Card Audio**: Mining a sentence card no longer fills the expression audio field; generated audio goes only to the configured sentence audio field.
- **Stats Mining Fields**: Sentence clips update `SentenceAudio` correctly; word audio uses configured Yomitan sources; English subtitle text is not written to word cards; secondary subtitle auto-selection prefers regular English tracks over Signs/Songs tracks.
- **Overlay Hover Readiness**: Subtitle bars are hoverable and clickable from the first subtitle line on visible overlay startup or resume, without waiting for the next subtitle event.
- **Startup Autoplay**: Playback is released after tokenization and overlay content are ready even when playback begins before the first subtitle line appears.
- **Overlay Startup Feedback**: Restored mpv OSD loading spinner that starts on connect, media open, or overlay request, and clears once the overlay is content-ready and visible.
- **Linux Overlay Input**: Notification close and action buttons remain clickable above subtitle bars on Linux.

<details>
<summary>Internal changes</summary>

**Internal**
- **Build**: `make deps` now initializes git submodules before installing dependencies on a fresh source checkout.
- **Release Tooling**: Release notes now credit contributors and first-time authors resolved from changelog fragments via git and the GitHub API.
- **Changelog Guidance**: PR fragment guidance updated to preserve separate-outcome fragments while directing contributors to consolidate same-PR follow-up notes before adding churn.

</details>

</details>

<details>
<summary>v0.15.x</summary>

<h2>v0.15.2 (2026-06-02)</h2>

**Changed**
- Yomitan: Updated the bundled Yomitan build to the latest vendored revision.

**Fixed**
- Anki - Animated AVIF: Clip timing no longer starts or ends early; word-audio lead-in and clip duration are now aligned to frame boundaries.
- Overlay (Hyprland): Fixed fullscreen overlay alignment - modal, stats, and sidebar content no longer shift below the mpv window.
- Overlay (macOS): Subtitle bars are now interactive immediately after autoplay starts with "wait for overlay to be ready" enabled, without requiring a manual click.
- Overlay (macOS): Fixed overlay, subtitles, and subtitle sidebar staying hidden after a modal closes until the user clicked the mpv window; focus is now restored to mpv when the last modal closes, so playback shortcuts and the overlay reappear correctly - including in native fullscreen.

<h2>v0.15.1 (2026-05-31)</h2>

**Fixed**

- **Linux Overlay Stacking**: Fixed the overlay intermittently dropping behind mpv on KDE Plasma and other non-Hyprland/Sway Wayland sessions; restored subtitle hover, pause-on-hover, and Yomitan lookups on X11/XWayland; the overlay now correctly layers above/below mpv based on fullscreen state, yields to foreground windows (Settings, Yomitan, AniList, etc.), and avoids startup flashes and fullscreen transition glitches.
- **Linux Overlay (Hyprland Lua)**: Fixed overlay placement on Hyprland 0.55+ when using a Lua-based config.
- **Manual Overlay Startup**: Fixed manual visible-overlay startup from mpv - now correctly attaches to playback, keeps the window bounds synced with mpv, and primes current subtitles before showing.
- **Playlist Transitions**: Reused the warm overlay when mpv advances to the next playlist item, avoiding a redundant tokenization pause and preserving visible subtitles across tracks.
- **macOS Overlay**: Fixed the visible subtitle overlay staying click-through after pause-until-ready releases playback; restored mpv focus after closing modal windows so subtitles and keybinds resume without clicking the player.
- **Mouse Keybindings**: Fixed keybinding capture and runtime handling for mouse buttons, including side buttons like `MBTN_BACK` and `MBTN_FORWARD`.
- **Windows mpv Shortcut**: Fixed the Windows `SubMiner mpv` shortcut so videos attach to an already-running background app instead of spawning a second process.

**Docs**

- **Troubleshooting**: Updated Hyprland overlay docs with current Lua (`hl.window_rule`) and legacy config syntax; added troubleshooting for KDE/Wayland and other non-Hyprland/Sway Wayland sessions; added a Character Dictionary troubleshooting section; added a "See Also" index linking each feature's troubleshooting page.

<h2>v0.15.0 (2026-05-29)</h2>

**Breaking Changes**

- **Subsync:**
  - The `subsync.defaultMode` config option has been removed
  - Subsync now always opens the manual subtitle picker regardless of any previously set default mode
- **N+1 Highlighting:**
  - N+1 highlighting now has its own dedicated `ankiConnect.nPlusOne.enabled` option, separate from known-word highlighting
  - It is no longer enabled automatically when known-word highlighting is on - enable it explicitly to keep N+1 annotations

**Added**

- **Auto-Updater:**
  - Tray and `subminer -u` update checks with app update prompts
  - Launcher and Linux rofi theme auto-updates
  - Checksum verification and configurable notifications
  - Opt-in prerelease channel via `updates.channel: "prerelease"`
- **Settings Window:**
  - New dedicated Settings window via `subminer --settings` or `subminer settings`, organized into Appearance, Behavior, Anki, Input, and Integration sections
  - Click-to-learn keybinding controls
  - AnkiConnect-backed deck, field, and note-type pickers that auto-fill from the configured Anki deck
  - Cross-category search
  - Live save for most options including subtitle CSS, stats keys, logging level, Jimaku, Subsync, and Anki mappings
  - AI and translation settings remain config-file only
- **Inline Character Portraits:**
  - Optional AniList character portraits appear inline for name-matched subtitle text
  - Manual AniList overrides scoped per parent media directory so separate season folders maintain separate character dictionary selections
- **Character Dictionary Manager:** New `Ctrl/Cmd+D` manager modal to remove, reorder, or override loaded entries.
- **Log Export:** Sanitized log ZIP export from the tray menu and via `subminer logs -e`, with home-directory usernames redacted from exported contents.
- **Launcher CLI:**
  - `subminer --version` / `subminer -v` prints the installed app version
  - `mpv.profile` config and Settings support passes a named mpv profile to managed launches
  - Bundled mpv plugin startup options are now configurable from SubMiner config
- **First-Run Setup:**
  - Optional installer for Bun and the `subminer` CLI on Linux, macOS, and Windows
  - Windows `subminer.cmd` PATH shim so `subminer` works without manually adding `SubMiner.exe` to PATH
  - Setup recognizes existing Homebrew or user PATH installs and avoids writing into Homebrew-owned paths
  - Includes an Open SubMiner Settings button
  - Standalone setup app quits after completing, returning terminal control
- **Primary Subtitle Visibility on Yomitan Popup:** New `subtitleStyle.primaryVisibleOnYomitanPopup` option keeps hover-mode primary subtitles visible while a Yomitan popup is open.

**Changed**

- **Subtitle Appearance Config:**
  - Primary and secondary subtitle appearance now use color controls plus CSS declaration editors, saved as `subtitleStyle.css`, `subtitleStyle.secondary.css`, and `subtitleSidebar.css`
  - Known-word and N+1 annotation colors moved to `subtitleStyle.knownWordColor` and `subtitleStyle.nPlusOneColor`
  - Subtitle font defaults updated to `Hiragino Sans, M PLUS 1, Source Han Sans JP, Noto Sans CJK JP`
  - Existing configs migrate automatically; legacy Anki color keys still accepted with deprecation warnings
- **Subtitle Style Defaults:**
  - Stronger outline-style text shadow
  - Thicker JLPT underlines
  - Frequency `topX` default raised to `10000`
- **Character Dictionary:**
  - Entries scoped to the current AniList media for name matching and inline portraits
  - Generates Japanese-only name aliases so raw romanized/English aliases no longer surface as separate results
  - In-app AniList selector waits for an explicit search with the box prefilled from the current filename
  - `subtitleStyle.nameMatchEnabled` is now the sole switch for dictionary sync and builds
- **Electron Runtime:** Updated from 39.8.6 to 42.2.0, returning SubMiner to a supported Electron release line.
- **Jellyfin Setup:**
  - Removed the server presets dropdown
  - Setup now shows a single editable server URL field
- **Jellyfin Cast Identity:**
  - Device identity now derived from the OS hostname and always reported as SubMiner
  - Previously configurable identity fields are ignored, preventing multiple installs from sharing a remote-session identity
- **Startup Defaults:** Jellyfin remote-session startup warmup and character-name subtitle highlighting now default to off.
- **Setup Appearance:** Removed the bundled mpv runtime plugin readiness card from the setup flow.

**Fixed**

- **AniList Progress:**
  - Progress updates fire correctly when playback reaches or skips past the watched threshold, using fresh mpv timing events
  - Season-specific results preferred for multi-season files, with a clear message when the matched season is not in Planning or Watching
  - Repeated missing-token checks no longer exhaust retry attempts or duplicate dead-letter entries
- **Anki Mining:**
  - Sentence-audio padding is opt-in by default
  - Animated AVIF freeze-frame duration aligned to word audio length without double-counting
  - Multi-line sentence alignment fixed for repeated subtitle text
  - Kiku duplicate-card detection, auto-merge, modal acknowledgment race, and field/tag ordering corrected
  - YouTube playback cards use mpv's resolved stream URLs
  - Sentence cards refresh the secondary subtitle before saving
- **Jellyfin Discovery:**
  - Startup, subtitle track selection, and duplicate ready-signal handling all fixed
  - Paused mpv no longer misreported as playing
  - Resume corrected when a remote play command sends `StartPositionTicks: 0` despite saved progress
- **Jellyfin Remote:**
  - Tray checkbox stays in sync on Linux after tray, CLI, or startup changes
  - Remote controller visibility and progress sync fixed for seeks, stops, startup path changes, and Linux websocket reconnect windows
  - Play and Resume now behave correctly (Play from beginning, Resume from saved position)
  - Final progress reports reuse SubMiner's last known position when mpv resets on stop
  - Windows setup login flow fixed with an IPC bridge, immediate feedback, and a timeout with inline error for unreachable servers
- **Overlay (macOS):**
  - Overlay hides when mpv loses focus, is minimized, or is no longer the foreground app
  - Stays stable through transient window geometry disappearances from macOS APIs and when clicking from the overlay back into mpv
  - Stats overlay opened inactive so it appears over fullscreen mpv without switching Spaces
  - Passthrough fixed so mpv controls stay clickable before hovering a subtitle bar
- **Yomitan Sidebar:**
  - Playback stays paused for sidebar-opened Yomitan popups when auto-pause is enabled
  - Popups now open when startup races the Yomitan extension load
  - Sidebar mining cards use audio and images from the clicked sidebar line instead of the current primary subtitle
- **Launcher:**
  - `subminer app` on Linux returns terminal control immediately
  - `subminer app --setup` opens the setup flow when SubMiner is already running in the background
- **YouTube Playback:**
  - Selected subtitles downloaded to local temp files so the primary bar and sidebar read the same source, with cleanup on reload and quit
  - False load-failure notifications suppressed
  - Tray icon created on launcher-managed playback that attaches to an already-running process
- **Shortcuts:**
  - Native mpv menu shortcuts disabled during managed macOS playback so configured SubMiner shortcuts work while mpv has focus
  - Custom session shortcuts including `stats.markWatchedKey` wired through mpv
  - Multi-line copy/mine overlay correctly focused so number keys choose the line count on macOS and Windows
- **Controller Bindings:**
  - Controller config and debug shortcuts stay closed while controller support is disabled
  - Binding learn mode starts from the edit pencil
  - Remaps saved per controller profile
  - Binding badges also start learn mode
  - Row reset buttons restore individual bindings to defaults
- **Logging:**
  - `logging.level` forwarded to launcher-started and Windows shortcut-started mpv sessions, covering mpv log verbosity, plugin logging, and plugin-launched app logging
  - `logging.rotation` (default 7 days) and per-component `logging.files` toggles added, with mpv logs disabled by default
  - Repeated IPC socket warning spam suppressed while waiting for mpv to recreate the socket
  - Windows mpv IPC, subtitle track, and Yomitan diagnostics added
- **In-Player Stats:**
  - Layering fixed so delete confirmations, overlay modals, and update-check dialogs appear above the stats window
  - Jellyfin playback stats grouped by item metadata so watched episodes merge with matching local library titles and keep clean display names
- **WebSocket Annotations:**
  - Annotation spans and token metadata stay on the annotation WebSocket
  - The regular subtitle WebSocket is plain-text only
- **Subtitle Annotation Prefetching:** Cached colored annotations and character images ready sooner for live subtitle changes without delaying raw subtitle display.
- **Windows Startup Errors:** Fatal startup failures now show a native error dialog and write details to the app log instead of exiting silently.

**Docs**

- **Documentation Site:**
  - Published stable docs at the site root with current development docs under `/main/`
  - Fixed versioned docs navigation, archived page link handling, and local dev version routing
  - Documented all previously undocumented config options including `subtitleStyle.primaryDefaultMode`, `stats.markWatchedKey`, `immersionTracking.lifetimeSummaries.*`, and all seven `mpv.*` launcher options
  - Added Playback Startup Flow and Runtime Sockets diagrams to the architecture docs with cross-reference pointers in the MPV Plugin and Troubleshooting pages

<details>
<summary>Internal changes</summary>

**Internal**

- **Release Tooling:**
  - Release-note polishing treats pending fragments and reviewed prerelease notes as a cumulative final outcome, collapsing prerelease-only fixes into the final user-facing change
  - Prerelease generation reuses existing reviewed notes and merges only new fragment material
  - `make clean` preserves `release/prerelease-notes.md`
- **Tests:** Removed stale Yomitan vendor source-inspection assertions for changes that were not shipped.

</details>

</details>

<details>
<summary>v0.14.x</summary>

<h2>v0.14.0 (2026-05-12)</h2>

**Added**

- **Character Dictionary:** Added AniList-based character dictionary selection for resolving title mismatches - open it in-app with the new `Ctrl+Alt+A` shortcut or from the CLI with `subminer dictionary --candidates` / `--select`. Series-scoped overrides replace stale entries in the merged dictionary.
- **Primary Subtitle Bar:** Added a `V` shortcut and mpv plugin binding to toggle the primary subtitle bar without affecting mpv's native subtitle visibility.
- **Texthooker:** Added `subminer texthooker -o` and a tray menu item to open the local texthooker page in the default browser.

**Changed**

- **mpv Plugin Setup:** SubMiner-managed launches now automatically inject the bundled mpv plugin when no global plugin is installed. The legacy plugin removal prompt appears before mpv starts so users can trash the old global plugin and relaunch immediately.
- **Tray Menu:** Replaced the "Open Overlay" menu item with "Open Help", which opens the session help modal.
- **Stats Exclusions:** Vocabulary exclusions now persist in the immersion database and import any existing browser-local exclusions on first load.
- **Config Example:** The generated example config now lists every built-in default keybinding, with regression coverage ensuring they compile, dispatch in the overlay, and register through the mpv plugin.
- **Default Startup Options:** Texthooker startup and subtitle/annotation WebSocket servers are now disabled by default. Fresh installs also get updated primary subtitle style defaults: a Japanese font stack, transparent subtitle backgrounds, stronger text shadows, and teal N4/fourth-band coloring.

**Fixed**

- **JLPT Underlines:** Restored persistent JLPT underlines on subtitle tokens; they now keep their JLPT color after dictionary lookups and when a token also carries a known-word or frequency annotation color.
- **Hyprland Fullscreen Overlay:** Fixed fullscreen transitions on Hyprland so mpv fullscreen changes refresh overlay geometry, reassert topmost stacking, and keep hover-pause working through resize/toggle cycles. The overlay now aligns exactly with the mpv window, disables floating-window decoration on exact placement, and no longer pins across workspaces.
- **macOS Overlay:** Kept the overlay visible and interactive during transient tracker refreshes and while mpv is the active tracked window. Fixed z-order so the overlay stays above mpv but behind other foreground windows.
- **Subtitle Annotation Colors:** Fixed annotated token colors so `subtitleStyle` typography is preserved and higher-priority known-word and frequency colors correctly take precedence over JLPT colors. Added a subtle brightness lift for hover states so transparent hover backgrounds still show a visible affordance. Hid the browser focus ring on the overlay surface so focused overlays no longer show a viewport border.
- **Grammar and Particle Annotation Filters:** Suppressed annotation styling (N+1, JLPT, frequency, name) for grammar-only tokens including auxiliary inflection fragments like `れる`/`れた`, polite copula tails like `です`/`ですよ`, negative copula phrases like `じゃないですか`, and standalone particles like `に`; lexical forms like `くれる` remain eligible.
- **Interjection and Existence Verb Filters:** Suppressed annotations for ハァ-style interjections and `ある`/`有る` existence verbs; known-word highlighting still applies for `ある`.
- **Known-Word Compound Tokens:** Preserved Yomitan compound token boundaries for known-word highlighting so known component words no longer color a larger unknown compound green.
- **Kana Token Annotations:** Stopped kana-only tokens from being selected as N+1 targets or receiving N+1, JLPT, or frequency annotation metadata.
- **Subtitle Bar Modes:** Changed `v` to cycle the primary subtitle bar through visible, hover, and hidden modes with OSD feedback. A new `subtitleStyle.primaryDefaultMode` config option sets the startup default independently of secondary subtitles.
- **Subtitle Keybindings:** Fixed the default replay/next subtitle keybindings by moving session help to `Ctrl/Cmd+/`, freeing `Ctrl+Shift+H` and `Ctrl+Shift+L` for subtitle playback controls. The mpv plugin now registers shifted letter chords correctly so `Ctrl+Shift+L` reaches play-next-subtitle; play-next also starts playback from a paused state before re-pausing at subtitle end.
- **Yomitan Popup Precedence:** Fixed keyboard-only Yomitan popup controls so popup shortcuts take precedence over overlay keybindings like `j`.
- **Subtitle Prefetch:** Kept annotation prefetch running after immediate cache-hit renders so upcoming subtitle colors stay ready.
- **Frequency Highlighting:** Fixed frequency highlighting for ordinal prefix-noun tokens like `第二` so JPDB ranks are preserved in subtitle annotations.
- **Known-Word Refresh After Mining:** Refreshed the current subtitle after successful card mining so newly known words recolor immediately.
- **Linux Multi-Line Copy:** Fixed multi-line subtitle copy on Linux so the overlay handles the follow-up digit selection locally, eliminating the timeout.
- **mpv Crash Notifications:** Stopped mpv from owning long-running SubMiner subprocesses during shutdown, preventing desktop crash notifications when closing video.
- **mpv Buffering Reload:** Kept the overlay alive across same-media reloads during buffering, avoiding duplicate startup gates and AniSkip lookups.
- **mpv Playlist Navigation:** Playlist navigation now reuses the running overlay without repeating the pause-until-ready warmup gate.
- **Anki Manual Updates:** Manual clipboard subtitle updates now replace sentence audio and media even when audio overwrite is disabled, while preserving existing word audio.
- **AniList Progress Tracking:** Fixed post-watch progress to check on mpv time-position updates using the fresh mpv position, wired manual mark-watched to force a progress sync, and filled missing episode metadata from the filename parser. Prevented duplicate post-watch writes during concurrent checks and preserved manual watched marks when sync fails.
- **AniList Setup:** Prevented config reload from opening the setup window during playback when token storage is unavailable; stopped setup from reporting success when encrypted token persistence fails.
- **AniList Linux Keyring:** Retried keyring availability after transient GNOME libsecret startup failures so stored tokens load once the keyring becomes available.
- **Stats Daemon:** Fixed stats background mode to route through the isolated stats daemon and deferred in-app stats startup when a daemon is already running, so video launches no longer fail when the stats port is in use.
- **Stats Session Detail:** Fixed recent session detail pages showing "Media not found" before lifetime media summaries are available.
- **Hover Background Config:** Accepted `subtitleStyle.hoverBackground` as an alias for `subtitleStyle.hoverTokenBackgroundColor` so setting it to `transparent` removes hover token backgrounds.
- **Managed Playback Exit:** Launcher-managed playback now exits the background SubMiner app when the video closes; explicit background launches remain persistent.
- **Multi-Mine Shortcuts:** Multi-line subtitle mining now accepts follow-up number-row digits even when the original shortcut modifiers are still held.
- **Jellyfin Setup:** Improved setup with recent server selection and inline authentication feedback; added a tray "Jellyfin Discovery" toggle for runtime-only cast discovery.
- **Subtitle POS Filtering:** Used Yomitan `wordClasses` metadata for subtitle part-of-speech filtering and backfilled blank MeCab POS detail fields during parser enrichment.

**Docs**

- Improved the docs homepage with canonical URLs and a cleaner sitemap for better search engine indexing.

<details>
<summary>Internal changes</summary>

**Internal**

- Replaced the changelog renderer with an AI polish pass that merges related fragments and writes user-facing release notes. `CHANGELOG.md` keeps internal items in a collapsed `<details>` block; GitHub release notes omit them entirely.
- Release CI no longer auto-builds pending `changes/*.md` fragments on tag. Tagging now fails fast if fragments remain - run `bun run changelog:build` (requires the `claude` CLI) and commit before tagging.

</details>

</details>

<details>
<summary>v0.12.x</summary>

<h2>v0.12.0 (2026-04-11)</h2>

**Changed**

- Overlay: Added configurable overlay shortcuts for session help, controller select, and controller debug actions.
- Overlay: Added mpv/plugin and CLI routing for session help, controller utilities, and subtitle sidebar toggling through the shared session-action path.
- Overlay: Improved dedicated overlay modal retry and focus handling for runtime options, Jimaku, session help, controller tools, and the playlist browser.
- Overlay: Fixed controller configuration and controller debug shortcut opens so configured bindings bring up their modals again instead of tripping renderer recovery.
- Stats: Sessions are rolled up per episode within each day, with a bulk delete that wipes every session in the group.
- Stats: Trends add a 365-day range next to the existing 7d/30d/90d/all options.
- Stats: Library detail view gets a delete-episode action that removes the video and all its sessions.
- Stats: Vocabulary Top 50 tightens the word/reading column so katakana entries no longer push the scores off screen.
- Stats: Episode detail hides card events whose Anki notes have been deleted, instead of showing phantom mining activity.
- Stats: Trend and watch-time charts share a unified theme with horizontal gridlines and larger ticks for legibility.
- Stats: Overview, Library, Trends, Sessions, and Vocabulary now use generic "title" wording so YouTube videos and anime live comfortably side by side in the dashboard.
- Stats: Session timeline no longer plots seek-forward/seek-backward markers - they were too noisy on sessions with lots of rewinds.
- Stats: Replaced the "Library - Per Day" section on the Stats → Trends page with a "Library - Summary" section. The new section shows a top-10 watch-time leaderboard chart and a sortable per-title table (watch time, videos, sessions, cards, words, lookups, lookups/100w, date range), all scoped to the current date range selector.

**Fixed**

- Overlay: Fixed overlay drag-and-drop routing so dropping external subtitle files like `.ass` onto mpv still loads them when the overlay is visible.
- Overlay: Addressed the latest CodeRabbit follow-ups on PR #49, including generation-scoped Lua session binding names, stricter session command validation, session-help shortcut visibility, the numeric-selection key guard, stats-overlay startup classification, and safer session-binding persistence.
- Overlay: Addressed the latest CodeRabbit follow-ups on the Windows overlay flow, including exact mpv target resolution, lower-overlay helper arguments, Win32 failure detection, and overlay cleanup on tracker loss.
- Overlay: Fixed Windows overlay z-order so the visible subtitle overlay stops staying above unrelated apps after mpv loses focus.
- Overlay: Fixed Windows overlay tracking to use native window polling and owner/z-order binding, which keeps the subtitle overlay aligned to the active mpv window more reliably.
- Overlay: Fixed Windows overlay hide/restore behavior so minimizing mpv immediately hides the overlay and restoring mpv brings it back on top of the mpv window without requiring a click.
- Overlay: Fixed stats overlay layering so the in-player stats page now stays above mpv and the subtitle overlay while it is open.
- Overlay: Fixed Windows subtitle overlay stability so transient tracker misses and restore events keep the current subtitle visible instead of waiting for the next subtitle line.
- Overlay: Fixed Windows focus handoff from the interactive subtitle overlay back to mpv so the overlay no longer drops behind mpv and briefly disappears.
- Overlay: Fixed Windows visible-overlay startup so it no longer briefly opens as an interactive or opaque surface before the tracked transparent overlay state settles.
- Overlay: Fixed spurious auto-pause after overlay visibility recovery and window resize so the overlay no longer pauses mpv until the pointer genuinely re-enters the subtitle area.
- Overlay: Fixed Windows secondary subtitle hover mode so the expanded hover hit area no longer blocks the native minimize, maximize, and close buttons.
- Overlay: Fixed Windows Yomitan popup focus loss after closing nested lookups so the original popup stays interactive instead of falling through to mpv.
- Stats: Fixed immersion-tracker timestamp handling under Bun/libsql so library rows, session timelines, and lifetime summaries keep real wall-clock millisecond values instead of truncating to invalid negative timestamps.
- Mpv Plugin: Fixed the mpv Lua plugin so hover and environment modules no longer use the `goto continue` pattern that can fail to parse on some user Lua runtimes.

**Internal**

- Release: Added a dedicated beta/rc prerelease GitHub Actions workflow that publishes GitHub prereleases without consuming pending changelog fragments or updating AUR.
- Release: Added prerelease note generation so beta and release-candidate tags can reuse the current pending `changes/*.md` fragments while leaving stable changelog publication for the final release cut.

</details>

<details>
<summary>v0.11.x</summary>

<h2>v0.11.2 (2026-04-07)</h2>

**Changed**

- Launcher: Replaced the launcher-only fullscreen toggle with `mpv.launchMode` so SubMiner-managed mpv playback can start in normal, maximized, or fullscreen mode.

**Fixed**

- Launcher: Fixed launcher-managed mpv spawning to force an explicit X11 GPU path when Wayland trackers are unavailable.
- Launcher: Local playback now promotes a single unlabeled external subtitle sidecar to the primary slot instead of leaving mpv's embedded English auto-selection in place.
- Release: Fixed Linux AppImage startup packaging so Chromium child relaunches can resolve the bundled `libffmpeg.so` instead of crash-looping on startup.

<h2>v0.11.1 (2026-04-04)</h2>

**Fixed**

- Release: Linux packaged builds now expose the canonical `SubMiner` app identity to Electron's startup metadata so native Wayland compositors stop reporting the window class/app-id as lowercase `subminer`.
- Linux: Linux now restores the runtime options, Jimaku, and Subsync shortcuts after the Electron 39 regression by routing those actions through the overlay's mpv/IPC shortcut path.

<h2>v0.11.0 (2026-04-03)</h2>

**Added**

- Overlay: Added a playlist browser overlay modal for browsing sibling video files and the live mpv queue during playback.
- Overlay: Added the default `Ctrl+Alt+P` keybinding to open the playlist browser and manage queue order without leaving playback.

**Changed**

- Setup: Made mpv plugin installation mandatory in the first-run setup flow, removed the skip path, and kept Finish disabled until the plugin is installed.
- Setup: Clarified that the mpv plugin requirement applies to setup on every platform, while the optional `SubMiner mpv` shortcut remains the recommended Windows playback entry point.
- Launcher: Streamlined Windows setup and config by making the `SubMiner mpv` shortcut self-contained and keeping `mpv.executablePath` as the simple fallback when `mpv.exe` is not on `PATH`.
- Overlay: Changed fresh-install default config to keep texthooker and stats from auto-opening browser tabs.
- Overlay: Changed fresh-install default config to enable AnkiConnect, Discord Rich Presence, subtitle-sidebar, and Yomitan-popup auto-pause by default, while disabling controller input by default.

**Fixed**

- Main: Resolve the YouTube playback socket path lazily so startup honors CLI and config overrides.
- Main: Add regression coverage for the lazy socket-path lookup during Windows mpv startup.
- Main: Keep integrated `--start --texthooker` launches on the full app-ready startup path so the texthooker page and websocket servers start together during normal playback startup.
- Main: Stop the mpv/plugin auto-start flow from spawning a separate standalone texthooker helper during normal `subminer <video>` launches.
- Overlay: Keep tracked macOS visible overlays click-through by default so subtitle sidebar passthrough works immediately without requiring a subtitle hover cycle first.
- Overlay: Add regression coverage for the macOS visible-overlay passthrough default.
- Anilist: Stop AniList post-watch from sending a second progress update when the current episode was already satisfied by a ready retry item in the same watch-completion pass.
- Anilist: Add regression coverage for the retry-queue plus live-update duplicate path.
- Overlay: Fixed Kiku duplicate grouping to reuse duplicate note IDs from both generic sentence-card creation and Yomitan popup mining instead of running extra duplicate scans after add.
- Overlay: Fixed the Yomitan popup mining flow to add cards in the background while keeping the stock popup progress feedback, then pause playback and close the lookup popup before the Kiku merge modal opens.
- Overlay: Fixed configured subtitle-jump keybindings so backward and forward subtitle seeks keep playback paused when invoked from a paused state.
- Launcher: Fixed the Windows `SubMiner mpv` shortcut and `SubMiner.exe --launch-mpv` flow to launch mpv with SubMiner's required default args directly instead of requiring an `mpv.conf` profile named `subminer`.
- Launcher: Clarified the Windows install and usage docs so the shortcut path is documented as self-contained, while the optional `subminer` mpv profile remains available for manual mpv launches.
- Launcher: Hardened the first-run setup blocker copy and stale custom-scheme handling so setup messages stay aligned with config, plugin, and dictionary readiness.
- Launcher: Fixed the Windows `SubMiner mpv` shortcut idle launch so loading a video after opening the shortcut keeps mpv in the expected SubMiner-managed session, auto-starts the overlay, and re-arms subtitle auto-selection for the newly opened file.
- Launcher: Removed the redundant `.` subtitle search path from the Windows shortcut launch args and deduped repeated subtitle source tracks in the manual sync picker so duplicate external subtitle entries no longer appear from the shortcut path.
- Playback: Fixed managed local playback so duplicate startup-ready retries no longer unpause media after a later manual pause on the same file.
- Playback: Fixed managed local subtitle auto-selection so local files reuse configured primary and secondary subtitle language priorities instead of staying on mpv's initial `sid=auto` guess.
- Launcher: Added a blank-by-default `mpv.executablePath` override for Windows playback so users can point SubMiner at `mpv.exe` when it is not on `PATH`.
- Launcher: Kept the Windows shortcut and `--launch-mpv` flow simple by preserving PATH auto-discovery as the default and exposing the override in first-run setup.
- Launcher: Added `windows` as a recognized launcher backend option and auto-detection target on Windows.
- Launcher: Honored `SUBMINER_YTDLP_BIN` consistently across YouTube playback URL resolution, track probing, subtitle downloads, and metadata probing.
- Launcher: Kept the first-run setup window from navigating away on unexpected URLs.
- Launcher: Made Windows mpv honor an explicitly configured executable path instead of silently falling back to PATH.
- Launcher: Hardened `--launch-mpv` parsing and Windows binary resolution so valueless flags do not swallow media targets and symlinked launcher installs do not short-circuit PATH lookup.
- Launcher: Fixed first-run setup blocking playback on macOS when the SubMiner mpv plugin was already installed at the canonical `~/.config/mpv` path.
- Launcher: Fixed setup gating so stale cancelled setup state no longer prevents playback when the canonical mpv plugin entrypoint already exists.
- Playback: Prevented stale async playlist-browser subtitle rearm callbacks from overriding newer subtitle selections during rapid file changes.

**Docs**

- Docs Site: Added a dedicated Subtitle Sidebar guide and linked it from the homepage and configuration docs.
- Docs Site: Linked Jimaku integration from the homepage to its dedicated docs page.
- Docs Site: Refreshed docs-site theme tokens and hover/selection styling for the updated pages.

**Internal**

- Release: Retried AUR clone and push operations in the tagged release workflow.
- Release: Kept GitHub Releases green when AUR publish flakes and needs manual follow-up.
- Release: Updated Electron to 39.8.6 and pinned patched transitive build dependencies to clear the reported high-severity audit findings.

</details>

<details>
<summary>v0.10.x</summary>

<h2>v0.10.0 (2026-03-29)</h2>

**Changed**

- Integrations: Replaced the deprecated Discord Rich Presence wrapper with the maintained `@xhayper/discord-rpc` package.

**Fixed**

- Stats: Fixed stats startup so the immersion tracker can run when `Bun.serve` is unavailable.
- Stats: Stats server now falls back to a Node `http` listener in Electron/runtime paths that do not expose Bun.
- Overlay: Fixed the macOS visible-overlay toggle path so manual hides stay hidden and the plugin uses the explicit visible-overlay toggle command.
- Subtitle Sidebar: Restored macOS mpv passthrough while the overlay subtitle sidebar is open so clicks outside the sidebar can refocus mpv and keep native keybindings working.

**Internal**

- Release: Added a maintained source coverage lane that shards Bun coverage one test file at a time and merges LCOV output into `coverage/test-src/lcov.info`.
- Release: CI and release quality-gate now upload the merged source-lane LCOV artifact for inspection.
- Runtime: Extracted remaining inline runtime logic from `src/main.ts` into dedicated runtime modules and composer helpers.
- Runtime: Added focused regression tests for the extracted runtime/composer boundaries.
- Runtime: Updated task tracking notes to mark TASK-238.6 complete and confirm follow-on boot-phase split can be deferred.
- Runtime: Split `src/main.ts` boot wiring into dedicated `src/main/boot/services.ts`, `src/main/boot/runtimes.ts`, and `src/main/boot/handlers.ts` modules.
- Runtime: Added focused tests for the new boot-phase seams and kept the startup/typecheck/build verification lanes green.
- Runtime: Updated internal architecture/task docs to record the boot-phase split and new ownership boundary.

</details>

<details>
<summary>v0.9.x</summary>

<h2>v0.9.3 (2026-03-25)</h2>

**Changed**

- Launcher: Moved YouTube primary subtitle language defaults to `youtube.primarySubLanguages`.
- Launcher: Removed the placeholder YouTube subtitle retime step and now uses downloaded primary subtitle tracks directly, so there is no fake path rewrite before playback/sidebar loading.
- YouTube: Removed the `src/core/services/youtube/retime` helper and its tests after retiring the internal retime strategy.
- Docs: Clarified optional `alass` / `ffsubsync` subtitle-sync requirements and setup steps, including fallback behavior when sync tools are absent.
- Launcher: Removed the old `youtubeSubgen.primarySubLanguages` config path from the generated config and docs.

<h2>v0.9.2 (2026-03-25)</h2>

**Fixed**

- Overlay: Fixed overlay pointer tracking so Windows click-through toggles immediately when the cursor enters or leaves subtitle regions, without waiting for a later hover resync.
- Overlay: Fixed Windows overlay window tracking on scaled displays by converting native tracked window bounds to Electron DIP coordinates before applying overlay bounds.
- Launcher: Fixed Windows direct `--youtube-play` startup so MPV boots reliably, stays paused until the app-owned subtitle flow is ready, and reuses an already-running SubMiner instance when available.
- Launcher: Fixed standalone Windows `--youtube-play` sessions so closing MPV fully exits SubMiner instead of leaving hidden overlay windows or a background process behind.
- Overlay: Fixed `subminer <youtube-url>` on Linux so the YouTube playback flow waits for Yomitan to load before creating the overlay window, avoiding the broken lookup popup state that previously required a manual overlay refresh.

<h2>v0.9.1 (2026-03-24)</h2>

**Changed**

- Release: Reduced packaged release size by excluding duplicate `extraResources` payload and pruning docs, tests, sourcemaps, and other source-only files from Electron bundles.

**Fixed**

- Overlay: Restored controller navigation and lookup/mining controls while the subtitle sidebar is open, while keeping true modal dialogs blocking controller actions.
- Tokenizer: Fixed subtitle annotation clearing so explanatory contrast endings like `んですけど` are excluded consistently across the shared tokenizer filter and annotation stage.

<h2>v0.9.0 (2026-03-23)</h2>

**Added**

- Docs: Added a new WebSocket / Texthooker API and integration guide covering WebSocket payloads, custom client patterns, mpv plugin automation, and webhook-style relay examples. Linked from configuration and mining workflow docs for easier discovery.

**Changed**

- Launcher: Added an app-owned YouTube subtitle flow that pauses mpv, uses absPlayer-style YouTube timedtext parsing/conversion to download subtitle tracks, and injects them as external files before playback resumes.
- Launcher: Changed YouTube subtitle startup to auto-load the best-available primary and secondary subtitle tracks at launch instead of forcing the picker modal first. Secondary subtitle failures no longer block playback resume.
- Launcher: Added `Ctrl+Alt+C` as the default keybinding to manually open the YouTube subtitle picker during active YouTube playback.
- Launcher: Added yt-dlp metadata probing so YouTube playback and immersion tracking record canonical video title and channel metadata.
- Launcher: Stopped forcing `--ytdl-raw-options=` before user-provided mpv options so existing YouTube cookie integrations in user `--args` are no longer clobbered.
- Launcher: Disabled mpv native YouTube subtitle auto-loading for the app-owned flow so injected external subtitle files remain authoritative.
- Launcher: Added OSD status messages for YouTube playback startup, subtitle acquisition, and subtitle loading so the flow stays visible before and during the picker.
- Subtitle Sidebar: Added startup-auto-open controls and resume positioning improvements so the sidebar jumps directly to the first resolved active cue.
- Subtitle Sidebar: Improved subtitle prefetch and embedded overlay passthrough sync so sidebar and overlay subtitle states stay consistent across media transitions.
- Subtitle Sidebar: Updated scroll handling, embedded layout styling, and active-cue visual behavior.
- Stats: Stats Library tab now displays YouTube video title, channel name, and channel thumbnail for YouTube media entries, with retry logic to fill in metadata that arrives after initial load.

**Fixed**

- Launcher: Fixed Anki media mining for mpv YouTube streams by unwrapping the stream URL so audio and screenshot capture work correctly for YouTube playback sessions.
- Immersion: Fixed YouTube media path handling in the immersion runtime and tracking so YouTube sessions record correct media references, AniList guessing skips YouTube URLs, and post-watch state transitions do not fire for YouTube media.
- Launcher: Fixed startup-launched YouTube playback so primary subtitle overlay updates continue after auto-load completes.
- Launcher: Fixed auto-loaded YouTube primary subtitles so parsed cues appear in the subtitle sidebar without needing a manual picker retry.
- Launcher: Fixed the YouTube picker to guard against duplicate subtitle submissions and tightened YouTube URL detection so follow-up runtime flows only treat real YouTube hosts as YouTube playback.
- Launcher: Fixed primary subtitle failure notifications being shown while app-owned YouTube subtitle probing and downloads are still in flight.
- Launcher: Preserved existing authoritative YouTube subtitle tracks when available; downloaded tracks are used only to fill missing sides, and native mpv secondary subtitle rendering is hidden so the overlay remains the sole secondary display.

</details>

<details>
<summary>v0.8.x</summary>

<h2>v0.8.0 (2026-03-22)</h2>

**Added**

- Overlay: Added the subtitle sidebar feature with a new `subtitleSidebar` configuration surface and rendered sidebar modal with cue list rendering, click-to-seek, active-cue highlighting, and embedded layout support.
- IPC: Added sidebar snapshot plumbing between renderer and main process for overlay/sidebar synchronization.

**Changed**

- Config: Added hot-reloadable sidebar options for enablement, layout, visibility, typography, opacity, sizing, and interaction behavior (`autoOpen`, `pauseOnHover`, `autoScroll`, toggle key).
- Docs: Added full `subtitleSidebar` documentation coverage, including sample config, option table, and toggle shortcut notes.
- Runtime: Improved subtitle prefetch/rendering flow so sidebar and overlay subtitle states stay in sync across media transitions.

**Fixed**

- Overlay: Kept sidebar cue tracking stable across playback transitions and timing edge cases.
- Overlay: Improved sidebar resume/start behavior to jump directly to the first resolved active cue.
- Overlay: Stopped stale subtitle refreshes from regressing active-cue and text state.

</details>

<details>
<summary>v0.7.x</summary>

<h2>v0.7.0 (2026-03-19)</h2>

**Added**

- Immersion: Added Mine Word, Mine Sentence, and Mine Audio buttons to word detail example lines in the stats dashboard.
- Immersion: Mine Word creates a full Yomitan card (definition, reading, pitch accent) via the hidden search page bridge, then enriches with sentence audio, screenshot, and metadata extracted from the source video.
- Immersion: Mine Sentence and Mine Audio create cards directly with appropriate Lapis/Kiku flags, sentence highlighting, and media from the source file.
- Immersion: Media generation (audio + image/AVIF) runs in parallel and respects all AnkiConnect config options.
- Immersion: Added word exclusion list to the Vocabulary tab with localStorage persistence and a management modal.
- Immersion: Fixed truncated readings in the frequency rank table (e.g. お前 now shows おまえ instead of まえ).
- Immersion: Clicking a bar in the Top Repeated Words chart now opens the word detail panel.
- Immersion: Secondary subtitle text is now stored alongside primary subtitle lines for use as translation when mining cards from the stats page.
- Stats: Added `subminer stats -b` to start or reuse a dedicated background stats server without blocking normal SubMiner instances.
- Stats: Added `subminer stats -s` to stop the dedicated background stats server without closing browser tabs.
- Stats: Stats server startup now reuses a running background stats daemon instead of trying to bind a second local server in another SubMiner instance.
- Launcher: Added launcher passthrough for `-a/--args` so mpv receives raw extra launch flags (`--fs`, `--ytdl-format`, custom audio/video settings, etc.) from the `subminer` command.
- Launcher: Added `subminer stats` to launch the local stats dashboard, force-start the stats server on demand, and open the dashboard in your browser.
- Launcher: Added `subminer stats cleanup` to backfill vocabulary metadata and prune stale or excluded immersion rows on demand.
- Launcher: Added `stats.autoOpenBrowser` so browser launch after `subminer stats` can be enabled or disabled explicitly.
- Immersion: Added a local stats dashboard for immersion tracking with Overview, Anime, Trends, Vocabulary, and Sessions views.
- Immersion: Added anime progress, episode completion, Anki card links, and occurrence drill-down across the stats dashboard.
- Immersion: Added richer session timelines with new-word activity, cumulative totals, and pause/seek/card event markers.
- Immersion: Added completed-episodes and completed-anime totals to the Overview tracking snapshot.

**Changed**

- Anki: Changed known-word cache settings to live under `ankiConnect.knownWords` instead of mixing them into `ankiConnect.nPlusOne`.
- Anki: Kept legacy `ankiConnect.nPlusOne` known-word keys and older `ankiConnect.behavior.nPlusOne*` keys as deprecated compatibility fallbacks.
- Stats: Added session deletion to the Sessions tab with the same confirmation prompt used by anime episode/session deletes, and removed all associated session rows from the stats database.
- Immersion: Kept immersion tracking history by default while preserving daily/monthly rollup maintenance.
- Immersion: Added exact lifetime summary reads for overview/anime/media stats so dashboard totals no longer depend on rescanning raw telemetry.
- Immersion: Reduced tracker storage overhead by removing duplicated subtitle text from subtitle-line event payloads.
- Immersion: Deduplicated episode cover-art blobs through a shared blob store and updated cover-art reads/writes to resolve shared images correctly.
- Immersion: Added indexes for large-history session, telemetry, vocabulary, kanji, and cover-art queries to keep dashboard reads fast as the SQLite database grows.
- Immersion: Renamed the stats dashboard's Anime tab to Library so the media browser label matches non-anime sources like YouTube and other yt-dlp-backed content.
- Anilist: Standardized episode completion threshold by introducing `DEFAULT_MIN_WATCH_RATIO` and using it for both local watched state transitions and AniList post-watch progress updates.
- Anilist: Episode auto-marking now uses the same threshold as AniList (`85%`), removing divergent completion behavior.
- Overlay: Excluded interjections and sound-effect tokens from subtitle annotation styling so they no longer inherit misleading lexical highlight treatment while still remaining visible and hoverable as plain subtitle tokens.
- Overlay: Expanded subtitle annotation noise filtering to also strip annotation metadata from standalone grammar-only helper tokens such as particles, auxiliaries, adnominals, common explanatory endings like `んです` / `のだ`, and merged trailing quote-particle forms like `...って` while keeping them tokenized for hover lookup.

**Fixed**

- Launcher: Fixed mpv Lua plugin binary auto-detection on Linux to also search `/usr/bin/subminer` and `/usr/local/bin/subminer` (lowercase), matching the conventional Unix wrapper name used by packaged installs such as the AUR package.
- Stats: Fixed the in-app stats overlay so it connects to the configured `stats.serverPort` instead of falling back to the default port.
- Overlay: Fixed subtitle frequency tagging for merged lookup-backed tokens like `陰に` by falling back to exact surface-form Yomitan frequencies when the normalized headword lookup misses.
- Overlay: Fixed MeCab merged-token position mapping across line breaks so merged content-plus-particle tokens like `陰に` keep their matched Yomitan frequency instead of inheriting shifted POS tags.
- Overlay: Fixed grouped frequency parsing in both Yomitan and fallback frequency-dictionary lookups so display values like `118,121` use the leading rank instead of collapsing the rank and occurrence count into `118121`.
- Overlay: Fixed frequency-rank ingestion to ignore Yomitan dictionaries explicitly marked `occurrence-based`, so raw occurrence counts are no longer treated as subtitle rank values.
- Overlay: Fixed inflected headword frequency tagging to prefer ranks from the selected Yomitan `termsFind` popup entry itself, ordered by configured dictionary priority, so forms like `潜み` use primary-dictionary ranks like `4073` before falling back to lower-priority raw lemma metadata such as `CC100`.
- Overlay: Fixed annotation-stage frequency filtering so exact kanji noun tokens like `者` keep their matched rank even when MeCab labels them `名詞/非自立`, instead of dropping the highlight after scan-time frequency lookup succeeds.
- Anki: Fixed repeated character-dictionary startup work by scheduling auto-sync only from mpv media-path changes instead of also re-triggering it from connection and media-title events for the same title.
- Overlay: Fixed macOS fullscreen overlay stability by keeping the passive visible overlay from stealing focus, re-raising the overlay window when reasserting its macOS topmost level, and tolerating one transient macOS tracker/helper miss before hiding the overlay.
- Overlay: Kept subtitle tokenization warmup one-shot for the lifetime of the app so later fullscreen/media churn on macOS does not replay the startup warmup gate after the first file is ready.
- Overlay: Added a bounded macOS tracker loss-grace window so fullscreen enter/leave transitions do not immediately hide and reload the overlay when the helper briefly loses the mpv window.
- Overlay: Skipped subtitle/tokenization refresh invalidation on character-dictionary auto-sync completion when the dictionary was already current, preventing startup flash/reload loops on unchanged media.
- Stats: Fixed session stats so known-word counts track real known-word occurrences without collapsing subtitle-line gaps.
- Stats: Fixed session word totals in session-facing stats views to prefer token counts when available, preventing known words from exceeding total words in the session chart.
- Stats: Fixed the stats Vocabulary tab blank-screen regression caused by a hook-order crash after vocabulary data finished loading.
- Anki: Fixed card-mine OSD feedback so the final mine result stops the Anki spinner first, then shows a single-line `✓`/`x` status without being overwritten by a later spinner tick.
- Stats: Removed the misleading `New words` series from expanded session charts; session detail now shows only the real total-word and known-word lines.
- Stats: Restored the cross-anime word table behavior in stats vocabulary surfaces so shared vocabulary entries no longer disappear or merge incorrectly across related media.
- Stats: `subminer stats -b` now runs as a standalone background stats daemon instead of reusing the main SubMiner app process, so the overlay app can still be launched separately for normal video watching.
- Stats: Dashboard word mining still works against the background daemon by using a short-lived hidden helper for the Yomitan add-note flow.
- Stats: Load full session timelines by default in stats session detail views so long sessions preserve complete telemetry history instead of being truncated by a fixed sample limit.
- Stats: Replaced heuristic stats word counts with Yomitan token counts, so session, media, anime, and trend subtitle totals now come directly from parsed subtitle tokens.
- Stats: Updated stats UI labels and lookup-rate copy to refer to tokens instead of words where those counts are shown.
- Overlay: Reduced repeated `Overlay loading...` popups on macOS when fullscreen tracker flaps briefly hide and recover the visible overlay.
- Stats: Scaled expanded session-detail known-word charts to the session's actual percentage range so small changes no longer render as a nearly flat line.
- Jlpt: Reduced JLPT dictionary startup log noise by summarizing duplicate surface-form collisions instead of logging one line per duplicate entry.

</details>

<details>
<summary>v0.6.x</summary>

<h2>v0.6.5 (2026-03-15)</h2>

**Internal**

- Release: Seed the AUR checkout with the repo `.SRCINFO` template before rewriting metadata so tagged releases do not depend on prior AUR state.

<h2>v0.6.4 (2026-03-15)</h2>

**Internal**

- Release: Reworked AUR metadata generation to update `.SRCINFO` directly instead of depending on runner `makepkg`, fixing tagged release publishing for `subminer-bin`.

<h2>v0.6.3 (2026-03-15)</h2>

**Changed**

- Overlay: Expanded the `Alt+C` controller modal into an inline config/remap flow with preferred-controller saving and per-action learn mode for buttons, triggers, and stick directions.

**Internal**

- Workflow: Hardened the `subminer-scrum-master` skill to explicitly answer whether docs updates and changelog fragments are required before handoff.
- Release: Automate `subminer-bin` AUR package updates from the tagged release workflow.

<h2>v0.6.2 (2026-03-12)</h2>

**Changed**

- Config: Added `yomitan.externalProfilePath` to reuse another Electron app's Yomitan profile in read-only mode.
- Config: SubMiner now reuses external Yomitan dictionaries/settings without writing back to that profile.
- Config: Launcher-managed playback now respects `yomitan.externalProfilePath` and no longer forces first-run setup when external Yomitan is configured.
- Config: SubMiner now seeds `config.jsonc` even when the default config directory already exists.
- Config: First-run setup now allows zero internal dictionaries when `yomitan.externalProfilePath` is configured, and falls back to requiring at least one internal dictionary if that external profile is later removed.

<h2>v0.6.1 (2026-03-12)</h2>

**Added**

- Overlay: Added Chrome Gamepad API controller support for keyboard-only overlay mode, including configurable logical bindings for lookup, mining, popup navigation, Yomitan audio, mpv pause, d-pad fallback navigation, and slower smooth popup scrolling.
- Overlay: Added `Alt+C` controller selection and `Alt+Shift+C` controller debug modals, with preferred controller persistence and live raw input inspection.
- Overlay: Added a transient in-overlay controller-detected indicator when a controller is first found.
- Overlay: Fixed stale keyboard-only token highlight cleanup when keyboard-only mode turns off or the Yomitan popup closes.

**Docs**

- Install: Added Arch Linux AUR install docs for `subminer-bin` in the README and installation guide.

**Internal**

- Config: add an enforced `verify:config-example` gate so checked-in example config artifacts cannot drift silently
- Release: Fixed the release workflow token permissions so tagged builds can download `oven-sh/setup-bun` and publish artifacts again.

</details>

<details>
<summary>v0.5.x</summary>

<h2>v0.5.6 (2026-03-10)</h2>

**Fixed**

- Dictionary: Persist merged character-dictionary MRU state as soon as a new retained set is built so revisits do not get dropped if later Yomitan import work fails, and skip merged dictionary rebuilds for reorder-only revisits when the retained anime set itself has not changed.
- Startup: Fixed early Electron startup writing config and user data under a lowercase `~/.config/subminer` path instead of the canonical `~/.config/SubMiner` directory.
- Overlay: Kept JLPT underline colors stable during Yomitan hover and selection states, even when tokens also use known, N+1, name-match, or frequency styling.

<h2>v0.5.5 (2026-03-09)</h2>

**Changed**

- Overlay: Added `f` as the default overlay fullscreen toggle and changed the default AniSkip intro-jump key to `Tab`.
- Dictionary: Aligned AniList character dictionary generation more closely with the upstream reference by preserving duplicate shared names across characters, skipping characters without native Japanese names, restoring richer character info fields, and using upstream-style role mapping plus hint-aware kanji readings.
- Startup: Ordered startup OSD messages so tokenization loads first, annotation loading appears next if still pending, and character dictionary sync progress waits until annotation loading finishes.
- Dictionary: Added a visible startup OSD step for merged character-dictionary building so long rebuilds show progress before the later import/upload phase.

**Fixed**

- Dictionary: Fixed AniList media guessing for character dictionary auto-sync by using filename-only `guessit` input and preserving multi-part guessit titles instead of truncating them to the first segment.
- Dictionary: Refresh the current subtitle after character dictionary auto-sync completes so newly imported character names highlight on the active line instead of waiting for the next subtitle change.
- Dictionary: Show character dictionary auto-sync progress on the mpv OSD without sending desktop notifications.
- Dictionary: Keep character dictionary auto-sync non-blocking during startup by letting snapshot/build work run in parallel and delaying only the Yomitan import/settings phase until current-media tokenization is already ready.
- Overlay: Fixed visible overlay keyboard handling so pressing `Tab` still reaches mpv and triggers the default AniSkip skip-intro binding while the overlay has focus.
- Plugin: Fix Windows mpv plugin binary override lookup so `SUBMINER_BINARY_PATH` still resolves to `SubMiner.exe` when no AppImage override is set.

<h2>v0.5.3 (2026-03-09)</h2>

**Changed**

- Release: Publish unsigned Windows `.exe` and `.zip` artifacts directly from release CI instead of routing them through SignPath.
- Release: Added `bun run build:win:unsigned` for explicit local unsigned Windows packaging.

<h2>v0.5.2 (2026-03-09)</h2>

**Internal**

- Release: Pinned the Windows SignPath submission workflow to an explicit artifact-configuration slug instead of relying on the SignPath project's default configuration.

<h2>v0.5.1 (2026-03-09)</h2>

**Changed**

- Launcher: Removed the YouTube subtitle generation mode switch so YouTube playback always preloads subtitles before mpv starts.

**Fixed**

- Launcher: Hardened YouTube AI subtitle fixing so fenced SRT output and text-only one-cue-per-block responses can still be applied without losing original cue timing.
- Launcher: Skipped AniSkip lookup during URL playback and YouTube subtitle-preload playback, limiting AniSkip to local file targets where it can actually resolve anime metadata.
- Launcher: Keep the background SubMiner process running after a launcher-managed mpv session exits so the next mpv instance can reconnect without restarting the app.
- Launcher: Reuse prior tokenization readiness after the background app is already warm so reopening a video does not pause again waiting for duplicate warmup completion.
- Windows: Acquire the app single-instance lock earlier so Windows overlay/video launches reuse the running background SubMiner process instead of booting a second full app and repeating startup warmups.

</details>

<details>
<summary>v0.3.x</summary>

<h2>v0.3.0 (2026-03-05)</h2>

- Added keyboard-driven Yomitan navigation and popup controls, including optional auto-pause.
- Added subtitle/jump keyboard handling fixes for smoother subtitle playback control.
- Improved Anki/Yomitan reliability with stronger Yomitan proxy syncing and safer extension refresh logic.
- Added Subsync `replace` option and deterministic retime naming for subtitle workflows.
- Moved aniskip resolution to launcher-script options for better control.
- Tuned tokenizer frequency highlighting filters for improved term visibility.
- Added release build quality-of-life for CLI publish (`gh`-based clobber upload).
- Removed docs Plausible integration and cleaned associated tracker settings.

</details>

<details>
<summary>v0.2.x</summary>

<h2>v0.2.3 (2026-03-02)</h2>

- Added performance and tokenization optimizations (faster warmup, persistent MeCab usage, reduced enrichment lookups).
- Added subtitle controls for no-jump delay shifts.
- Improved subtitle highlight logic with priority and reliability fixes.
- Fixed plugin loading behavior to keep OSD visible during startup.
- Fixed Jellyfin remote resume behavior and improved autoplay/tokenization interaction.
- Updated startup flow to load dictionaries asynchronously and unblock first tokenization sooner.

<h2>v0.2.2 (2026-03-01)</h2>

- Improved subtitle highlighting reliability for frequency modes.
- Fixed Jellyfin misc info formatting cleanup.
- Version bump maintenance for 0.2.2.

<h2>v0.2.1 (2026-03-01)</h2>

- Delivered Jellyfin and Subsync fixes from release patch cycle.
- Version bump maintenance for 0.2.1.

<h2>v0.2.0 (2026-03-01)</h2>

- Added task-related release work for the overlay 2.0 cycle.
- Introduced Overlay 2.0.
- Improved release automation reliability.

</details>

<details>
<summary>v0.1.x</summary>

<h2>v0.1.2 (2026-02-24)</h2>

- Added encrypted AniList token handling and default GNOME keyring support.
- Added launcher passthrough for password-store flows (Jellyfin path).
- Updated docs for auth and integration behavior.
- Version bump maintenance for 0.1.2.

<h2>v0.1.1 (2026-02-23)</h2>

- Fixed overlay modal focus handling (`grab input`) behavior.
- Version bump maintenance for 0.1.1.

<h2>v0.1.0 (2026-02-23)</h2>

- Bootstrapped Electron runtime, services, and composition model.
- Added runtime asset packaging and dependency vendoring.
- Added project docs baseline, setup guides, architecture notes, and submodule/runtime assets.
- Added CI release job dependency ordering fixes before launcher build.

</details>
