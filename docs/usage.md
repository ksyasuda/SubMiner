# Usage

There are two ways to use SubMiner — the `subminer` wrapper script or the mpv plugin:

| Approach            | Best For                                                                                                                                                                                                       |
| ------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **subminer script** | All-in-one solution. Handles video selection, launches MPV with the correct socket, and manages app commands. Overlay start is explicit (`--start`, `-S`, or `y-s`).                                           |
| **MPV plugin**      | When you launch MPV yourself or from other tools. Provides in-MPV chord keybindings (e.g. `y-y` for menu) to control visible and invisible overlay layers. Requires `--input-ipc-server=/tmp/subminer-socket`. |

You can use both together—install the plugin for on-demand control, but use `subminer` when you want the streamlined workflow.

`subminer` is implemented as a Bun script and runs directly via shebang (no `bun run` needed), for example: `subminer --start video.mkv`.

## Live Config Reload

While SubMiner is running, it watches your active config file and applies safe updates automatically.

Live-updated settings:

- `subtitleStyle`
- `keybindings`
- `shortcuts`
- `secondarySub.defaultMode`
- `ankiConnect.ai`

Invalid config edits are rejected; SubMiner keeps the previous valid runtime config and shows an error notification.
For restart-required sections, SubMiner shows a restart-needed notification.

## Commands

```bash
# Browse and play videos
subminer                          # Current directory (uses fzf)
subminer -R                       # Use rofi instead of fzf
subminer -d ~/Videos              # Specific directory
subminer -r -d ~/Anime            # Recursive search
subminer video.mkv                # Play specific file
subminer --start video.mkv        # Play + explicitly start overlay
subminer https://youtu.be/...     # Play a YouTube URL
subminer ytsearch:"jp news"       # Play first YouTube search result
subminer --log-level debug video.mkv # Enable verbose logs for launch/debugging
subminer --log-level warn video.mkv  # Set logging level explicitly

# Options
subminer -T video.mkv             # Disable texthooker server
subminer -b x11 video.mkv         # Force X11 backend
subminer video.mkv                # Uses mpv profile "subminer" by default
subminer -p gpu-hq video.mkv      # Override mpv profile
subminer jellyfin                 # Open Jellyfin setup window (subcommand form)
subminer jellyfin -l --server http://127.0.0.1:8096 --username me --password 'secret'
subminer jellyfin --logout        # Clear stored Jellyfin token/session data
subminer jellyfin -p              # Interactive Jellyfin library/item picker + playback
subminer jellyfin -d              # Jellyfin cast-discovery mode (foreground app)
subminer doctor                   # Dependency + config + socket diagnostics
subminer config path              # Print active config path
subminer config show              # Print active config contents
subminer mpv socket               # Print active mpv socket path
subminer mpv status               # Exit 0 if socket is ready, else exit 1
subminer mpv idle                 # Launch detached idle mpv with SubMiner defaults
subminer texthooker               # Launch texthooker-only mode
subminer app --anilist            # Pass args directly to SubMiner binary (example: AniList login flow)
subminer yt -o ~/subs https://youtu.be/...  # YouTube subcommand: output directory shortcut
subminer yt --mode preprocess --whisper-bin /path/to/whisper-cli --whisper-model /path/to/model.bin https://youtu.be/...  # Pre-generate subtitle tracks before playback

# Direct AppImage control
SubMiner.AppImage --background             # Start in background (tray + IPC wait, minimal logs)
SubMiner.AppImage --start --texthooker   # Start overlay with texthooker
SubMiner.AppImage --texthooker           # Launch texthooker only (no overlay window)
SubMiner.AppImage --stop                  # Stop overlay
SubMiner.AppImage --start --toggle        # Start MPV IPC + toggle visibility
SubMiner.AppImage --start --toggle-invisible-overlay  # Start MPV IPC + toggle invisible layer
SubMiner.AppImage --show-visible-overlay              # Force show visible overlay
SubMiner.AppImage --hide-visible-overlay              # Force hide visible overlay
SubMiner.AppImage --show-invisible-overlay            # Force show invisible overlay
SubMiner.AppImage --hide-invisible-overlay            # Force hide invisible overlay
SubMiner.AppImage --start --dev                         # Enable app/dev mode only
SubMiner.AppImage --start --debug                       # Alias for --dev
SubMiner.AppImage --start --log-level debug             # Force verbose logging without app/dev mode
SubMiner.AppImage --settings              # Open Yomitan settings
SubMiner.AppImage --jellyfin              # Open Jellyfin setup window
SubMiner.AppImage --jellyfin-login --jellyfin-server http://127.0.0.1:8096 --jellyfin-username me --jellyfin-password 'secret'
SubMiner.AppImage --jellyfin-logout       # Clear stored Jellyfin token/session data
SubMiner.AppImage --jellyfin-libraries
SubMiner.AppImage --jellyfin-items --jellyfin-library-id LIBRARY_ID --jellyfin-search anime --jellyfin-limit 20
SubMiner.AppImage --jellyfin-play --jellyfin-item-id ITEM_ID --jellyfin-audio-stream-index 1 --jellyfin-subtitle-stream-index 2  # Requires connected mpv IPC (--start or plugin workflow)
SubMiner.AppImage --jellyfin-remote-announce  # Force cast-target capability announce + visibility check
SubMiner.AppImage --help                  # Show all options
```

### Logging and App Mode

- `--log-level` controls logger verbosity.
- `--dev` and `--debug` are app/dev-mode switches; they are not log-level aliases.
- `--background` defaults to quieter logging (`warn`) unless `--log-level` is set.
- `--background` launched from a terminal detaches and returns the prompt; stop it with tray Quit or `SubMiner.AppImage --stop`.
- Linux desktop launcher starts SubMiner with `--background` by default (via electron-builder `linux.executableArgs`).
- On Linux, the app now defaults `safeStorage` to `gnome-libsecret` for encrypted token persistence.
  Override with e.g. `--password-store=basic_text`.
- Use both when needed, for example `SubMiner.AppImage --start --dev --log-level debug`.

### Launcher Subcommands

- `subminer jellyfin` / `subminer jf`: Jellyfin-focused workflow aliases.
- `subminer yt` / `subminer youtube`: YouTube-focused shorthand flags (`-o`, `-m`).
- `subminer doctor`: health checks for core dependencies and runtime paths.
- `subminer config`: config helpers (`path`, `show`).
- `subminer mpv`: mpv helpers (`status`, `socket`, `idle`).
- `subminer texthooker`: texthooker-only shortcut (same behavior as `--texthooker`).
- `subminer app` / `subminer bin`: direct passthrough to the SubMiner binary/AppImage.
- Subcommand help pages are available (for example `subminer jellyfin -h`, `subminer yt -h`).

Use subcommands for Jellyfin/YouTube command families (`subminer jellyfin ...`, `subminer yt ...`).
Top-level launcher flags like `--jellyfin-*` and `--yt-subgen-*` are intentionally rejected.

### MPV Profile Example (mpv.conf)

`subminer` passes the following MPV options directly on launch by default:

- `--input-ipc-server=/tmp/subminer-socket` (or your configured socket path)
- `--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us`
- `--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us`
- `--sub-auto=fuzzy`
- `--sub-file-paths=.;subs;subtitles`
- `--sid=auto`
- `--secondary-sid=auto`
- `--secondary-sub-visibility=no`

You can define a matching profile in `~/.config/mpv/mpv.conf` for consistency when launching `mpv` manually or from other tools. `subminer` launches with `--profile=subminer` by default (or override with `subminer -p <profile> ...`):

```ini
[subminer]
# IPC socket (must match SubMiner config)
input-ipc-server=/tmp/subminer-socket

# Prefer JP/EN audio + subtitle language variants
alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us
slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us

# Auto-load external subtitles
sub-auto=fuzzy
sub-file-paths=.;subs;subtitles

# Select primary + secondary subtitle tracks automatically
sid=auto
secondary-sid=auto
secondary-sub-visibility=no
```

`secondary-slang` is not an mpv option; use `slang` with `sid=auto` / `secondary-sid=auto` instead.

### YouTube Playback

`subminer` accepts direct URLs (for example, YouTube links) and `ytsearch:` targets, and forwards them to mpv.

Notes:

- Install `yt-dlp` so mpv can resolve YouTube streams and subtitle tracks reliably.
- `subminer` supports three subtitle-generation modes for YouTube URLs:
  - `automatic` (default): starts playback immediately, generates subtitles in the background, and loads them into mpv when ready.
  - `preprocess`: generates subtitles first, then starts playback with generated `.srt` files attached.
  - `off`: disables launcher generation and leaves subtitle handling to mpv/yt-dlp.
- Primary subtitle target languages come from `youtubeSubgen.primarySubLanguages` (defaults to `["ja","jpn"]`).
- Secondary target languages come from `secondarySub.secondarySubLanguages` (defaults to English if unset).
- `subminer` prefers subtitle tracks from yt-dlp first, then falls back to local `whisper.cpp` (`whisper-cli`) when tracks are missing.
- Whisper translation fallback currently only supports English secondary targets; non-English secondary targets rely on yt-dlp subtitle availability.
- Configure defaults in `$XDG_CONFIG_HOME/SubMiner/config.jsonc` (or `~/.config/SubMiner/config.jsonc`) under `youtubeSubgen` and `secondarySub`, or override mode/tool paths via CLI flags/environment variables.

## Keybindings

### Global Shortcuts

| Keybind       | Action                   |
| ------------- | ------------------------ |
| `Alt+Shift+O` | Toggle visible overlay   |
| `Alt+Shift+I` | Toggle invisible overlay |
| `Alt+Shift+Y` | Open Yomitan settings    |

`Alt+Shift+Y` is a fixed global shortcut; it is not part of `shortcuts` config.

### Overlay Controls (Configurable)

| Input                | Action                                             |
| -------------------- | -------------------------------------------------- |
| `Space`              | Toggle MPV pause                                   |
| `ArrowRight`         | Seek forward 5 seconds                             |
| `ArrowLeft`          | Seek backward 5 seconds                            |
| `ArrowUp`            | Seek forward 60 seconds                            |
| `ArrowDown`          | Seek backward 60 seconds                           |
| `Shift+H`            | Jump to previous subtitle                          |
| `Shift+L`            | Jump to next subtitle                              |
| `Ctrl+Shift+H`       | Replay current subtitle (play to end, then pause)  |
| `Ctrl+Shift+L`       | Play next subtitle (jump, play to end, then pause) |
| `Q`                  | Quit mpv                                           |
| `Ctrl+W`             | Quit mpv                                           |
| `Right-click`        | Toggle MPV pause (outside subtitle area)           |
| `Right-click + drag` | Move subtitle position (on subtitle)               |
| `Ctrl/Cmd+Shift+P`   | Toggle invisible subtitle position edit mode       |
| `Arrow keys`         | Move invisible subtitles while edit mode is active |
| `Enter` / `Ctrl+S`   | Save invisible subtitle position in edit mode      |
| `Esc`                | Cancel invisible subtitle position edit mode       |
| `Ctrl/Cmd+A`         | Append clipboard video path to MPV playlist        |

These keybindings only work when the overlay window has focus. See [Configuration](/configuration) for customization.

### Drag-and-drop Queueing

- Drag and drop one or more video files onto the overlay to replace current playback (`loadfile ... replace` for first file, then append remainder).
- Hold `Shift` while dropping to append all dropped files to the current MPV playlist.

## How It Works

1. MPV runs with an IPC socket at `/tmp/subminer-socket`
2. The overlay connects and subscribes to subtitle changes
3. Subtitles are tokenized with Yomitan's internal parser
4. Words are displayed as clickable spans
5. Clicking a word triggers Yomitan popup for dictionary lookup
6. Texthooker server runs at `http://127.0.0.1:5174` for external tools
