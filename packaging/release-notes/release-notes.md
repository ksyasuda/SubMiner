## Highlights
### Added
- **Hachidori Dictionary Backend**: Bundled Hachidori backend as an alternative to Yomitan. Select it with `dictionaryBackend` and restart SubMiner.
  - Works with subtitle scanning, popups, frequency highlighting, character dictionaries, Anki mining, and stats dashboard mining.
  - First-run setup builds Hachidori's Anki template from your existing SubMiner deck, tags, and field mappings.
  - Can link an external Hachidori app, browser, or Docker host while Anki mining stays in SubMiner.
  - The tray and dictionary-settings shortcut open the selected backend's settings. `--hachidori` and `--yomitan` open a specific one.
- **Jellyfin Subtitle Generation**: Japanese subtitle generation now works while playing Jellyfin episodes, with both direct-play and transcoded playback.

### Fixed
- **Media Timing Review**: Opens on the first try and keeps playback paused while open, even when a dictionary popup closes.
- **Jellyfin Terminal Playback**:
  - `subminer jellyfin -p` lists libraries when SubMiner is already running, and picker image previews work again.
  - Real errors, like an expired session or an unreachable server, now appear instead of an empty list.
  - Watch progress and played status reach Jellyfin even when cast discovery is off.
  - When it starts SubMiner itself, SubMiner now shows a tray icon and quits when mpv closes.
  - The rofi search step is now a compact input box.
- **macOS Casting**: Jellyfin casting finds mpv when SubMiner is opened from Finder, the Dock, or Spotlight. `mpv.executablePath` and `SUBMINER_MPV_PATH` are now honored on macOS and Linux.
- **macOS Launcher PATH**: The command-line launcher no longer shows "Not on PATH" when SubMiner is opened from Finder or the Dock.
- **Keyboard-Only Mode**: `h`, `l`, and the arrow keys type normally in modals and text fields instead of seeking mpv. Toggling the mode inside a modal no longer turns on a hidden keyboard-only mode.
- **Halfwidth Katakana**: Halfwidth katakana in subtitles now displays as fullwidth, including in mined cards and the stats dashboard. Stats sentence search matches katakana in either width.
- **Subtitle Annotations**: Lines shown while two cues overlap now get annotations right away. Prefetching no longer restarts during startup.
- **Subtitle Sidebar**: The open sidebar does far less work during playback.
- **Hyprland and Sway**: A slow or hung `hyprctl` or `swaymsg` no longer freezes the app.
- **Stats Mining**: Cards mined from the stats dashboard keep the selected history line and media.
- **Security Updates**: Updated Electron, undici, and follow-redirects to fix new security advisories.

### Docs
- **Hachidori Setup**: Covers backend selection, Hachidori dictionary and Anki setup, settings flags and shortcuts, and linking an external host.
- **Jellyfin**: The docs now say SubMiner joins the cast menu at startup only when both `startupWarmups.jellyfinRemoteSession` and `jellyfin.remoteControlAutoConnect` are on. Otherwise, enable **Jellyfin Discovery** in the tray.
- **Casting and mpv**: Added how casting finds mpv on each platform, a troubleshooting entry for casts that never start mpv, and a note on the macOS 27 Accessibility list rename.
- **Jimaku Environment Variables**: Removed the unused `SUBMINER_JIMAKU_*` variables from the docs. Set the key with `jimaku.apiKey` or `jimaku.apiKeyCommand` instead.

## What's Changed

- feat(dictionary): add Hachidori backend support by @ksyasuda in #266
- fix(jellyfin): make terminal play work with a running app and report progress by @ksyasuda in #288
- fix(overlay): stop Hyprland/Sway tracker polls from blocking the main thread by @ksyasuda in #289
- fix(overlay): poll cheap playback state in the open subtitle sidebar by @ksyasuda in #290
- fix(overlay): keep keyboard mode from swallowing modal and text-field keys by @ksyasuda in #291
- refactor: remove dead code and enable noUnusedLocals by @ksyasuda in #292
- feat(subtitles): generate Japanese subtitles during Jellyfin playback by @ksyasuda in #294
- fix(subtitles): widen halfwidth katakana where subtitle text enters the app by @ksyasuda in #295
- fix(jellyfin): find mpv when SubMiner is launched from Finder by @ksyasuda in #298

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
