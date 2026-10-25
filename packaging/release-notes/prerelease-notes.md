> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

<!-- prerelease-version: 0.20.1-beta.2; since: v0.20.1-beta.1 -->

## Changes since v0.20.1-beta.1

- Upgraded the desktop runtime to Electron 43.7.7.
- SubMiner now refuses to start on an unsupported Electron runtime, or after an Electron major or minor downgrade, so your Yomitan profile data stays safe; patch downgrades are still allowed.
- Automatic character-dictionary updates now stop if a Yomitan profile that had dictionaries suddenly reports none.
- Development launches now use a separate `SubMiner-dev` profile unless you explicitly ask for the production profile.
- Sentence mining (single and multi-line) no longer counts each card twice in session and lifetime stats.
- Deleting sessions now updates word and kanji "first seen" and "last seen" dates correctly, and retention pruning keeps historical discovery dates and counts.
- Words hidden by the vocabulary filter are no longer counted in the new-words charts, so the charts now match the vocabulary summary.
- A one-time background repair on next launch fixes stale "last seen" dates and existing chart mismatches.
- Overview totals, charts, the activity calendar, and the vocabulary summary now refresh after you delete sessions, and a failed refresh keeps the dashboard visible with a retry option.
- Subtitle generation now works during Windows YouTube playback and loads the generated subtitles into the playing video.
- On Windows, moving on to a video queued from the YouTube browser window no longer drops the rest of mpv's queue.
- The stats docs now cover the automatic Overview refresh after deletions and how discovery dates are kept through retention pruning.

## Highlights
### Added

- **YouTube Browser Window**:
  - Open it with `subminer youtube` (or `subminer yt`), or pick **Browse YouTube** from the tray. It stays logged in across restarts.
  - Clicking a video plays it in mpv. Middle-click, Shift-click, Ctrl-click, or the right-click menu adds it to mpv's playlist.
  - Queued videos get the same Japanese subtitle setup when they start.
  - Closing the window quits SubMiner if you started it with `subminer youtube`. If a video is still playing, SubMiner waits until mpv closes.
- **Whisper Subtitles for YouTube**:
  - Set `youtube.subtitleSource` to `whisper` to make subtitles with Whisper instead of downloading YouTube's captions.
  - The video stays paused while Whisper works, and progress shows in the subtitle generation modal. You can close the modal to keep watching, or cancel.
  - The subtitle generation modal (`Ctrl+Shift+G`) now works on YouTube videos, including on Windows.
- **Mouse and Scroll Wheel Bindings**:
  - You can now use scroll wheel keys (`WHEEL_UP`, `WHEEL_DOWN`, `WHEEL_LEFT`, `WHEEL_RIGHT`) in `keybindings` and record them in the settings key editor.
  - mpv's mouse button and wheel bindings now work while the cursor is over the overlay. This matters most on Hyprland, where these bindings did nothing before.
  - When a binding conflicts, SubMiner's own bindings and right-click pause take priority over mpv's.
- **Session Help Commands**: During playback, press `Enter` or double-click a row in session help to run that command. Hovering over a row makes it the target for `Enter`.

### Changed

- **Release Assets**: Releases no longer include package-size JSON reports.

### Fixed

- **Jellyfin Subtitles**:
  - The Japanese track is selected, and annotations start, as soon as it downloads. SubMiner no longer waits for the other tracks.
  - Japanese and English tracks are now picked automatically on episodes with image-based subtitles (PGS, DVD, DVB).
  - If one track fails to download, the other tracks are still selected.
- **Jellyfin Audio Track**: On direct-play streams, timing review previews and mined card audio now use the audio track you're playing.
- **Media Timing Review**:
  - Trim handles are now thin lines centered on the clip edges.
  - Original subtitle boundaries are now plain orange bars with no labels.
  - The preview cursor no longer stutters on macOS with Bluetooth headphones such as AirPods.
- **Hyprland Fullscreen**: The overlay stays visible when mpv enters or leaves fullscreen. With `mpv.backend: x11`, subtitles and the sidebar stay clickable after going fullscreen.
- **mpv Key Bindings in the Overlay**: mpv key bindings, such as `9`/`0` for volume, now work while the overlay has focus, even when the overlay loaded before mpv connected.
- **YouTube Japanese Subtitles**: Some videos also list a machine-translated Japanese track. On those, Japanese subtitles no longer fail with "HTTP 429". If YouTube refuses a direct download, SubMiner retries with yt-dlp.
- **Stats Accuracy**:
  - Sentence mining no longer counts each card twice in session and lifetime stats.
  - Overview totals, charts, the activity calendar, and the vocabulary summary now update after you delete sessions.
  - Word and kanji dates stay correct after deleting sessions. New-word charts now match the vocabulary summary.
- **Dictionary Safety**: Automatic character dictionary changes now stop if a Yomitan profile that had dictionaries suddenly reports none. SubMiner also refuses to load Yomitan data after a downgrade to an older Electron version.
- **AppImage in Sandboxes**: The AppImage now opens its window in sandboxes that mount the image themselves, such as `firejail --appimage`.
- **Ctrl+C Crash**: Stopping the terminal command that started SubMiner with Ctrl+C no longer crashes it with a "write EPIPE" error.
- **Security Updates**: Electron is updated to 43.7.7 and undici to 7.29.1, which fixes known security advisories.

### Docs

- **Docs Site Rewrite**: Pages are shorter and easier to scan. The configuration reference now has a short explanation and a table of keys and defaults for each config block. Pages that no longer matched current behavior are fixed.
- **Arch Linux Install**: The MeCab install command now uses the AUR package `mecab-git`. The README now matches the installation guide.
- **Stats**: The docs now cover the Overview refreshing after you delete sessions, and how first-seen dates are kept when old sessions are cleaned up.

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
