# SubMiner Script vs MPV Plugin

There are two ways to use SubMiner:

| Approach            | Best For                                                                                                                                                                              |
| ------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **subminer script** | All-in-one solution. Handles video selection, launches MPV with the correct socket, starts the overlay automatically, and cleans up on exit.                                          |
| **MPV plugin**      | When you launch MPV yourself or from other tools. Provides in-MPV chord keybindings (e.g. `y-y` for menu) to control visible and invisible overlay layers. Requires `--input-ipc-server=/tmp/subminer-socket`. |

You can use both together—install the plugin for on-demand control, but use `subminer` when you want the streamlined workflow.

`subminer` is implemented as a Bun script and runs directly via shebang (no `bun run` needed), for example: `subminer video.mkv`.

## Usage

```bash
# Browse and play videos
subminer                          # Current directory (uses fzf)
subminer -R                       # Use rofi instead of fzf
subminer -d ~/Videos              # Specific directory
subminer -r -d ~/Anime            # Recursive search
subminer video.mkv                # Play specific file
subminer https://youtu.be/...     # Play a YouTube URL
subminer ytsearch:"jp news"       # Play first YouTube search result

# Options
subminer -T video.mkv             # Disable texthooker server
subminer -b x11 video.mkv         # Force X11 backend
subminer video.mkv                # Uses mpv profile "subminer" by default
subminer -p gpu-hq video.mkv      # Override mpv profile
subminer --yt-subgen-mode preprocess --whisper-bin /path/to/whisper-cli --whisper-model /path/to/model.bin https://youtu.be/...  # Pre-generate subtitle tracks before playback

# Direct AppImage control
subminer.AppImage --start --texthooker   # Start overlay with texthooker
subminer.AppImage --texthooker           # Launch texthooker only (no overlay window)
subminer.AppImage --stop                  # Stop overlay
subminer.AppImage --start --toggle        # Start MPV IPC + toggle visibility
subminer.AppImage --start --toggle-invisible-overlay  # Start MPV IPC + toggle invisible layer
subminer.AppImage --show-visible-overlay              # Force show visible overlay
subminer.AppImage --hide-visible-overlay              # Force hide visible overlay
subminer.AppImage --show-invisible-overlay            # Force show invisible overlay
subminer.AppImage --hide-invisible-overlay            # Force hide invisible overlay
subminer.AppImage --settings              # Open Yomitan settings
subminer.AppImage --help                  # Show all options
```

### MPV Profile Example (mpv.conf)

Add a profile to `~/.config/mpv/mpv.conf`; `subminer` now launches mpv with `--profile=subminer` by default (or override with `subminer -p <profile> ...`):

```ini
[subminer]
# IPC socket (must match SubMiner config)
input-ipc-server=/tmp/subminer-socket

# Prefer JP subs, then EN
slang=ja,jpn,en,eng

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

| Keybind       | Action                    |
| ------------- | ------------------------- |
| `Alt+Shift+O` | Toggle visible overlay    |
| `Alt+Shift+I` | Toggle invisible overlay  |
| `Alt+Shift+Y` | Open Yomitan settings     |

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

These keybindings only work when the overlay window has focus. See [Configuration](/configuration) for customization.

## How It Works

1. MPV runs with an IPC socket at `/tmp/subminer-socket`
2. The overlay connects and subscribes to subtitle changes
3. Subtitles are tokenized with Yomitan's internal parser, with MeCab fallback when needed
4. Words are displayed as clickable spans
5. Clicking a word triggers Yomitan popup for dictionary lookup
6. Texthooker server runs at `http://127.0.0.1:5174` for external tools
