## Highlights
### Added
- **Plugins**:
  - JavaScript plugins load from `<config dir>/plugins/<id>/` and are enabled under `plugins` in the config.
  - Plugins can react to mpv events and mined cards, call AnkiConnect, run programs, and show results in an overlay panel or window.
  - Three examples are included: `ai-explain` for line explanations, `anki-sync` for AnkiWeb syncing, and `episode-tags` for tagging cards by show and episode.
- **Subtitle Filters**: Remove parenthesized text, sound descriptions, speaker names, music notes, or custom regex matches from primary subtitles. Configure them in Settings under Behavior → Subtitle Filters.
- **Controller Dialog Layouts**: The media timing review and duplicate card dialogs can now be fully operated by controller and show button hints. Remap both layouts per controller in `Alt+C`.
- **Controller Customization**:
  - `Alt+C` actions can be combos: click **Learn**, hold one button, and press another.
  - New unbound actions cover subtitle navigation, holding to show the secondary subtitle, Yomitan entry selection, and character navigation.
  - **Button Style** offers Xbox, Nintendo, or PlayStation names and symbols, saved per controller.
  - `Alt+C` now warns when two actions share an input.
- **Keyboard-Only Mode**: Press `C` to move and look up by character, reaching words split by subtitle tokenization. **Pause Video In Keyboard-Only Mode** can also pause when the mode turns on and resume when it turns off.
- **Jimaku Auto-Open**: Set `jimaku.autoOpenWhenPrimarySubtitleMissing` to pause and open Jimaku when a local video has no Japanese subtitle track. Playback resumes when the modal closes.

### Changed
- **Stats Trends**: Rebuilt with period comparisons, Learning charts, a heatmap, a title timeline, and a single running-total chart. Week grouping, custom ranges, a page-wide title filter, and click-through to sessions and library pages are also new.
- **Stats Vocabulary**:
  - Rebuilt around one searchable, sortable word table with Worth mining, All words, Across titles, and Newest views.
  - Added known words by frequency band, a vocabulary growth chart, and a full kanji view.
  - New word counts now count dictionary words, so a word's conjugations count once and totals are lower than before.
  - Words and lines from streams now say they can't be mined, instead of showing Unavailable buttons.
- **Controller Behavior**: **Toggle MPV Pause** now works outside keyboard-only mode. Controller connected and disabled messages now follow your regular notification setting instead of using a separate overlay pill.
- **Keyboard-Only Mode Exit**: Turning the mode off now also closes the dictionary popup.
- **Hachidori**: The bundled dictionary reader is updated to 0.4.0, adding every pitch accent in each popup entry, pitch drawn on the headword reading, per-kanji furigana, and Yomitan-style definition tags.

### Fixed
- **Controller Settings (`Alt+C`)**:
  - Remaps now apply right away, without restarting SubMiner.
  - Editing stays locked until you press a button, so the window can't edit a controller it can't see.
  - With several controllers connected, it selects and listens to the one you just pressed.
  - **Learn**, **Clear**, and **Reset** are reachable again, D-pad **Learn** no longer errors, and button names are correct on standard-layout pads.
- **Controller Input**: One press no longer runs its action twice on macOS and Windows. On macOS, buttons now work while mpv is in front.
- **Dictionary Popup**: Controller and keyboard lookups pause and resume playback reliably, and moving between words right after closing the popup no longer reopens it.
- **Anki Sentence Audio**: `ankiConnect.fields.audio` now defaults to `SentenceAudio`, so generated audio no longer overwrites Yomitan's word audio. If you relied on the old default, set it to `ExpressionAudio`.
- **AniList**: Progress sync now adds an exactly matched or pinned show to your Watching or Completed list when it isn't on your list yet. Completed, Paused, Dropped, and Rewatching entries are never changed.
- **Stats Dashboard**:
  - Overview, media, library, trends, vocabulary, and sentence search load several times faster on large databases.
  - Trends totals, episode counts, rate charts, axis labels, and date handling in time zones ahead of UTC are corrected.
  - Word lists now total every form of a word and show the dictionary form's reading.
  - Names no longer slip past Hide Names, and totals stay correct after vocabulary cleanup.
- **Playback**: Reloading the current file in mpv no longer leaves playback paused, and the "Overlay loading" message no longer gets stuck when moving to the next playlist file.
- **AppImage**: "Service Crash" notifications on quit no longer appear on Ubuntu 26.04 and newer.

### Docs
- **Keyboard and Controller Reference**: Keyboard shortcuts now list the keyboard-only `C` key, and the configuration reference points to Controller support for default button actions.

## What's Changed

- feat(plugins): add JavaScript plugin system with overlay panel by @ksyasuda in #303
- fix(anilist): add missing list entry on confident match by @dandascan in #305
- Controller: hold to show the secondary subtitle, dictionary entry navigation, standard button names by @YashaNoos in #312
- Controller: layouts for the media timing review and duplicate card dialogs by @YashaNoos in #313
- Controller: subtitle line actions, Button Style, keyboard-only mode pause, Alt+C fixes by @YashaNoos in #316
- fix(plugin): keep loading OSD off on warm playlist advance by @ksyasuda in #317
- feat(subtitles): add regex subtitle filters with built-in presets by @ksyasuda in #318
- feat(jimaku): auto-open modal when a video has no Japanese subtitle by @dandascan in #319
- fix(controller): pick up macOS presses while mpv is in front by @ksyasuda in #321
- perf(stats): speed up dashboard queries on large databases by @ksyasuda in #322
- fix(mpv): resume playback after a same-file reload that skips the path clear by @ksyasuda in #323
- fix(controller): keep popup pause and press handling reliable by @ksyasuda in #324
- feat(stats): overhaul the Trends tab by @ksyasuda in #325
- Controller: character navigation in keyboard-only mode, plain Xbox face names by @YashaNoos in #326
- feat(controller): add hold + press combos to main bindings by @ksyasuda in #328

## New Contributors

- @dandascan made their first contribution in #305
- @YashaNoos made their first contribution in #312

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
