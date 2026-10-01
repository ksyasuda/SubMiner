> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

<!-- prerelease-version: 0.20.1-beta.1 -->

## Highlights

### Added

- **YouTube Browser Window**:
  - Open a YouTube browser with `subminer youtube` (or `subminer yt`), or pick **Browse YouTube** from the tray. It stays logged in to YouTube across restarts.
  - Clicking a video plays it in mpv instead of on the page.
  - Middle-click, Shift-click, Ctrl-click, or the right-click menu adds a video to mpv's playlist.
  - When a queued video starts, it gets the same Japanese subtitle setup.
  - If you started SubMiner with `subminer youtube`, closing the YouTube window quits SubMiner. If a video is still playing, SubMiner quits after mpv closes.
- **Whisper Subtitles for YouTube**:
  - Set `youtube.subtitleSource` to `whisper` to make subtitles with Whisper instead of downloading YouTube's captions. The default is still `youtube`.
  - The video stays paused while Whisper works, and progress shows in the subtitle generation modal. Close the modal to keep watching, or cancel to carry on without subtitles.
  - The subtitle generation modal (`Ctrl+Shift+G`) now works on YouTube videos with either `youtube.subtitleSource` setting.
  - Whisper downloads only a small audio stream and deletes it when generation ends, or when you switch videos, close mpv, or quit.
- **Mouse and Scroll Wheel Bindings**:
  - `keybindings` now accepts scroll wheel keys (`WHEEL_UP`, `WHEEL_DOWN`, `WHEEL_LEFT`, `WHEEL_RIGHT`), with or without modifiers. You can also record them in the settings key editor.
  - Mouse button and scroll wheel bindings from mpv now work while the cursor is over the overlay. This covers your `input.conf` and mpv's defaults, such as double-click for fullscreen, the wheel for volume, and back/forward for the playlist.
  - Hyprland users benefit most: the overlay always takes input there, so these bindings did nothing before.
  - SubMiner's own bindings and right-click pause still win over mpv's.
- **Session Help Commands**:
  - During playback, press `Enter` or double-click a row in session help to run that command.
  - Hovering over a row makes it the target for `Enter`.
  - Session help opened from the tray with no video loaded is still read-only.

### Changed

- **Release Assets**: Releases no longer include package-size JSON reports. Package contents are still validated before release.

### Fixed

- **Jellyfin Subtitles**:
  - The Japanese subtitle track is selected, and annotations start, as soon as that track downloads. SubMiner no longer waits for every other track.
  - Subtitle tracks now download in parallel, so a slow embedded track no longer holds up the main subtitles.
  - Episodes with image-based subtitles (PGS, DVD, DVB) now select the Japanese and English tracks automatically.
  - If one subtitle track fails to download, the other tracks are still selected.
- **Jellyfin Audio Track**: On direct-play Jellyfin streams, timing review previews and mined card audio now use the audio track you're playing, not the one Jellyfin marks as default.
- **Media Timing Review**:
  - Trim handles are now thin lines centered on the clip edges. The start and end of the clip are easier to see, and less of the waveform is hidden.
  - The original subtitle boundaries are now plain orange bars without text labels, so the waveform stays visible.
  - On macOS with Bluetooth headphones such as AirPods, the preview cursor now moves smoothly. Before, it stuttered and jumped to the end partway through. This affected Opus audio, including Jellyfin streams.
- **Hyprland Fullscreen**:
  - The subtitle overlay now stays visible when mpv enters or leaves fullscreen, so focus changes no longer cancel fullscreen. Sway still refreshes the overlay as before.
  - With `mpv.backend: x11`, subtitles and the sidebar no longer stop responding to clicks after mpv goes fullscreen.
- **mpv Key Bindings in the Overlay**:
  - mpv key bindings, such as `9`/`0` for volume, now work while the overlay has focus.
  - They used to do nothing if the overlay loaded before SubMiner connected to mpv, which was common with `mpv.backend: x11`. The overlay now reloads mpv's bindings once mpv connects.
- **YouTube Japanese Subtitles**:
  - Some videos also list a machine-translated Japanese track. On those, Japanese subtitles no longer fail with "HTTP 429", because SubMiner now picks the real Japanese track.
  - If YouTube refuses a direct subtitle download, SubMiner retries with yt-dlp.
- **AppImage in Sandboxes**: The AppImage now opens its window when started in a sandbox that mounts the image itself, such as `firejail --appimage`. Before, it exited without opening a window.
- **Ctrl+C Crash**: Stopping the terminal command that started SubMiner with Ctrl+C no longer crashes SubMiner with a "write EPIPE" error.
- **Security Updates**: Electron is updated to 42.10.0 and undici to 7.29.1 to fix known security advisories.

### Docs

- **Docs Site Rewrite**:
  - Pages are shorter and easier to scan. Each one starts with setup and usage, and reference details are in compact tables.
  - The configuration reference now gives each config block a short explanation and a table of keys and defaults.
  - Pages that no longer matched how SubMiner works have been corrected.
- **Arch Linux Install**:
  - The MeCab install command is fixed. MeCab is only in the AUR, as `mecab-git`.
  - The README's requirements and quick start now match the installation guide.

## What's Changed

- fix(jellyfin): select subtitles faster and survive failed track downloads by @ksyasuda in #270
- fix(jellyfin): use the current audio track for media generation by @ksyasuda in #271
- feat(youtube): add YouTube browser window and Whisper subtitle source by @ksyasuda in #273
- fix(overlay): thin centered trim handles and label-free subtitle boundaries in timing review by @ksyasuda in #274
- feat(overlay): run session help commands with Enter or double-click by @ksyasuda in #275
- fix(overlay): keep overlay mapped during Hyprland/Sway fullscreen by @ksyasuda in #276
- fix(linux): skip AppImage mount keepalive on sandbox squashfs mounts by @ksyasuda in #277
- fix(overlay): reload imported mpv keys after mpv connects by @ksyasuda in #278
- feat(overlay): forward mpv mouse button and wheel bindings by @ksyasuda in #279

## Installation

See the [installation guide](https://docs.subminer.moe/installation) for full setup steps.

## Assets

- Linux: `SubMiner-*.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz`, the `subminer` launcher, and the Windows `subminer.cmd` launcher
- Bun corresponding source: `bun-v1.3.5-source.tar.gz` and its `.sha256` file

Both launcher downloads use Bun included with the SubMiner app. Download `subminer` on Linux or macOS and `subminer.cmd` on Windows.

The app bundles an unmodified Bun 1.3.5 runtime. Bun is MIT licensed and statically links JavaScriptCore (LGPL 2.0) and TinyCC (LGPL 2.1). License texts and third-party notices ship inside the app under `resources/bun/licenses`, and the source archive above contains the matching Bun, WebKit, and dependency sources for relinking.
