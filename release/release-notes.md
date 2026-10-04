## Highlights
### Added
- **YouTube Browser Window**:
  - Browse YouTube in a window that keeps your login across restarts. Open it with `subminer youtube` (or `subminer yt`), or choose **Browse YouTube** from the tray.
  - Clicking a video plays it in mpv. Middle-click, Shift/Ctrl-click, or the right-click menu adds it to mpv's playlist.
  - Queued videos get the same Japanese subtitle setup when they start playing.
- **Whisper Subtitles for YouTube**: Set `youtube.subtitleSource` to `whisper` to transcribe YouTube videos with Whisper instead of using YouTube's captions. The subtitle generation modal (`Ctrl+Shift+G`) now also works on YouTube videos.
- **Scroll Wheel and Mouse Bindings**: `keybindings` now accepts scroll wheel keys, and the settings key editor can capture them. Mouse button and wheel bindings from mpv now work while the cursor is over the overlay.
- **Run Commands From Session Help**: Press `Enter` or double-click a row in session help to run that command during playback.

### Fixed
- **YouTube**:
  - Japanese subtitles no longer fail with "HTTP 429" on videos that also have a machine-translated Japanese track.
  - On Windows, the subtitle generation modal now works during YouTube playback.
  - On Windows, playing a video queued from the YouTube browser window no longer clears the rest of mpv's queue.
- **Jellyfin**:
  - The Japanese subtitle track is selected as soon as it downloads, instead of waiting for every other track.
  - Japanese and English image-based subtitles (PGS, DVD, DVB) are now selected automatically.
  - Timing review previews and mined card audio use the audio track you are playing.
- **Hyprland Overlay**: The overlay no longer interrupts mpv when it goes fullscreen. With `mpv.backend: x11`, the overlay stays clickable after mpv goes fullscreen.
- **mpv Key Bindings**: Keys from mpv's defaults and your `input.conf` (such as `9` and `0` for volume) now work while the overlay has focus, even when the overlay loads before SubMiner connects to mpv.
- **Media Timing Review**: Trim handles are now thin lines on the clip edges, and the original subtitle boundaries are plain orange bars. On macOS with Bluetooth headphones such as AirPods, the preview cursor now moves smoothly.
- **Stats**:
  - Each mined card is counted once instead of twice.
  - Overview totals, charts, the calendar, and the vocabulary summary update right after you delete sessions.
  - Word and kanji dates stay correct after deletes and pruning, and the new-words charts now agree with the vocabulary summary. Existing data is repaired in the background on the next launch.
- **Yomitan Profile Safety**: Electron is updated to 43.7.7, and SubMiner now refuses to load your Yomitan data with an older or unsupported runtime. Automatic character-dictionary updates stop if a profile that had dictionaries suddenly reports none.
- **Security Updates**: Updated Electron and undici to resolve security advisories.
- **AppImage in Sandboxes**: The AppImage now opens a window when launched inside a sandbox such as `firejail --appimage`.
- **Ctrl+C Crash**: SubMiner no longer crashes with "write EPIPE" when you stop the terminal command that started it with Ctrl+C.

### Docs
- **Docs Site**: The docs site is shorter and easier to scan, and the configuration reference now has a table of keys and defaults for each section. Outdated pages were corrected, and the Arch Linux MeCab install now uses the AUR package `mecab-git`.
- **Stats**: Added docs on how the Overview refreshes after you delete sessions and how discovery dates are kept after pruning.

## What's Changed

- fix(storage): protect Yomitan profiles across Electron runtime changes by @ksyasuda in #211
- fix(jellyfin): select subtitles faster and survive failed track downloads by @ksyasuda in #270
- fix(jellyfin): use the current audio track for media generation by @ksyasuda in #271
- feat(youtube): add YouTube browser window and Whisper subtitle source by @ksyasuda in #273
- fix(overlay): thin centered trim handles and label-free subtitle boundaries in timing review by @ksyasuda in #274
- feat(overlay): run session help commands with Enter or double-click by @ksyasuda in #275
- fix(overlay): keep overlay mapped during Hyprland/Sway fullscreen by @ksyasuda in #276
- fix(linux): skip AppImage mount keepalive on sandbox squashfs mounts by @ksyasuda in #277
- fix(overlay): reload imported mpv keys after mpv connects by @ksyasuda in #278
- feat(overlay): forward mpv mouse button and wheel bindings by @ksyasuda in #279
- fix(stats): count mined cards once and keep lexical stats exact after deletes by @ksyasuda in #280
- fix(youtube): keep mpv queue when advancing to a queued video on Windows by @ksyasuda in #281
- fix(subtitles): allow Whisper generation during Windows YouTube stream playback by @ksyasuda in #282

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
