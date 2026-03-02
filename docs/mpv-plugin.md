# MPV Plugin

The SubMiner mpv plugin (`subminer/main.lua`) provides in-player keybindings to control the overlay without leaving mpv. It communicates with SubMiner by invoking the AppImage (or binary) with CLI flags.

## Installation

```bash
# From release bundle:
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer-assets.tar.gz -O /tmp/subminer-assets.tar.gz
tar -xzf /tmp/subminer-assets.tar.gz -C /tmp
mkdir -p ~/.config/SubMiner
cp /tmp/config.example.jsonc ~/.config/SubMiner/config.jsonc
mkdir -p ~/.config/mpv/scripts/subminer
mkdir -p ~/.config/mpv/script-opts
cp -R /tmp/plugin/subminer/. ~/.config/mpv/scripts/subminer/
cp /tmp/plugin/subminer.conf ~/.config/mpv/script-opts/

# Or from source checkout: make install-plugin
```

mpv must have IPC enabled for SubMiner to connect:

```ini
# ~/.config/mpv/mpv.conf
input-ipc-server=/tmp/subminer-socket
```

## Keybindings

All keybindings use a `y` chord prefix — press `y`, then the second key:

| Chord | Action                 |
| ----- | ---------------------- |
| `y-y` | Open menu              |
| `y-s` | Start overlay          |
| `y-S` | Stop overlay           |
| `y-t` | Toggle visible overlay |
| `y-o` | Open settings window   |
| `y-r` | Restart overlay        |
| `y-c` | Check status           |
| `y-k` | Skip intro (AniSkip)   |

## Menu

Press `y-y` to open an interactive menu in mpv's OSD:

```text
SubMiner:
1. Start overlay
2. Stop overlay
3. Toggle overlay
4. Open options
5. Restart overlay
6. Check status
```

Select an item by pressing its number.

## Configuration

Create or edit `~/.config/mpv/script-opts/subminer.conf`:

```ini
# Path to SubMiner binary. Leave empty for auto-detection.
binary_path=

# MPV IPC socket path. Must match input-ipc-server in mpv.conf.
socket_path=/tmp/subminer-socket

# Enable the texthooker WebSocket server.
texthooker_enabled=yes

# Port for the texthooker server.
texthooker_port=5174

# Window manager backend: auto, hyprland, sway, x11, macos.
backend=auto

# Start the overlay automatically when a file is loaded.
# Runs only when mpv input-ipc-server matches socket_path.
auto_start=yes

# Show the visible overlay on auto-start.
# Runs only when mpv input-ipc-server matches socket_path.
auto_start_visible_overlay=yes

# Pause mpv on visible auto-start until SubMiner signals overlay/tokenization readiness.
# Requires auto_start=yes and auto_start_visible_overlay=yes.
auto_start_pause_until_ready=yes

# Show OSD messages for overlay status changes.
osd_messages=yes

# Logging level: debug, info, warn, error.
log_level=info

# Enable AniSkip intro detection/markers.
aniskip_enabled=yes

# Optional title override (launcher fills from guessit when available).
aniskip_title=

# Optional season override (launcher fills from guessit when available).
aniskip_season=

# Optional MAL ID override. Leave blank to resolve from media title.
aniskip_mal_id=

# Optional episode override. Leave blank to detect from filename/title.
aniskip_episode=

# Show OSD skip button while inside intro range.
aniskip_show_button=yes

# OSD label + keybinding for intro skip action.
aniskip_button_text=You can skip by pressing %s
aniskip_button_key=y-k
aniskip_button_duration=3
```

### Option Reference

| Option                         | Default                       | Values                                     | Description                                                                      |
| ------------------------------ | ----------------------------- | ------------------------------------------ | -------------------------------------------------------------------------------- |
| `binary_path`                  | `""` (auto-detect)            | file path                                  | Path to SubMiner binary                                                          |
| `socket_path`                  | `/tmp/subminer-socket`        | file path                                  | MPV IPC socket path                                                              |
| `texthooker_enabled`           | `yes`                         | `yes` / `no`                               | Enable texthooker server                                                         |
| `texthooker_port`              | `5174`                        | 1–65535                                    | Texthooker server port                                                           |
| `backend`                      | `auto`                        | `auto`, `hyprland`, `sway`, `x11`, `macos` | Window manager backend                                                           |
| `auto_start`                   | `yes`                         | `yes` / `no`                               | Auto-start overlay on file load when mpv socket matches `socket_path`            |
| `auto_start_visible_overlay`   | `yes`                         | `yes` / `no`                               | Show visible layer on auto-start when mpv socket matches `socket_path`           |
| `auto_start_pause_until_ready` | `yes`                         | `yes` / `no`                               | Pause mpv on visible auto-start; resume when SubMiner signals tokenization-ready |
| `osd_messages`                 | `yes`                         | `yes` / `no`                               | Show OSD status messages                                                         |
| `log_level`                    | `info`                        | `debug`, `info`, `warn`, `error`           | Log verbosity                                                                    |
| `aniskip_enabled`              | `yes`                         | `yes` / `no`                               | Enable AniSkip intro detection                                                   |
| `aniskip_title`                | `""`                          | string                                     | Override title used for lookup                                                   |
| `aniskip_season`               | `""`                          | numeric season                             | Optional season hint                                                             |
| `aniskip_mal_id`               | `""`                          | numeric MAL id                             | Skip title lookup; use fixed id                                                  |
| `aniskip_episode`              | `""`                          | numeric episode                            | Skip episode parsing; use fixed                                                  |
| `aniskip_show_button`          | `yes`                         | `yes` / `no`                               | Show in-range intro skip prompt                                                  |
| `aniskip_button_text`          | `You can skip by pressing %s` | string                                     | OSD prompt format (`%s`=key)                                                     |
| `aniskip_button_key`           | `y-k`                         | mpv key chord                              | Primary key for intro skip action (`y-k` always works as fallback)               |
| `aniskip_button_duration`      | `3`                           | float seconds                              | OSD hint duration                                                                |

## Binary Auto-Detection

When `binary_path` is empty, the plugin searches platform-specific locations:

**Linux:**

1. `~/.local/bin/SubMiner.AppImage`
2. `/opt/SubMiner/SubMiner.AppImage`
3. `/usr/local/bin/SubMiner`
4. `/usr/bin/SubMiner`

**macOS:**

1. `/Applications/SubMiner.app/Contents/MacOS/SubMiner`
2. `~/Applications/SubMiner.app/Contents/MacOS/SubMiner`

**Windows:**

1. `C:\Program Files\SubMiner\SubMiner.exe`
2. `C:\Program Files (x86)\SubMiner\SubMiner.exe`
3. `C:\SubMiner\SubMiner.exe`

## Backend Detection

When `backend=auto`, the plugin detects the window manager:

1. **macOS** — detected via platform or `OSTYPE`.
2. **Hyprland** — detected via `HYPRLAND_INSTANCE_SIGNATURE`.
3. **Sway** — detected via `SWAYSOCK`.
4. **X11** — detected via `XDG_SESSION_TYPE=x11` or `DISPLAY`.
5. **Fallback** — defaults to X11 with a warning.

## Script Messages

The plugin can be controlled from other mpv scripts or the mpv command line using script messages:

```
script-message subminer-start
script-message subminer-stop
script-message subminer-toggle
script-message subminer-menu
script-message subminer-options
script-message subminer-restart
script-message subminer-status
script-message subminer-autoplay-ready
script-message subminer-aniskip-refresh
script-message subminer-skip-intro
```

The `subminer-start` message accepts overrides:

```
script-message subminer-start backend=hyprland socket=/custom/path texthooker=no log-level=debug
```

`log-level` here controls only logging verbosity passed to SubMiner.
`--debug` is a separate app/dev-mode flag in the main CLI and should not be used here for logging.

## AniSkip Intro Skip

- AniSkip lookups are gated. The plugin only runs lookup when:
  - SubMiner launcher metadata is present, or
  - SubMiner app process is already running, or
  - You explicitly call `script-message subminer-aniskip-refresh`.
- Lookups are asynchronous (no blocking `ps`/`curl` on `file-loaded`).
- MAL/title resolution is cached for the current mpv session.
- When launched via `subminer`, launcher runs `guessit` first (file targets) and passes title/season/episode to the plugin; fallback is filename-derived title.
- Install `guessit` for best detection quality (`python3 -m pip install --user guessit`).
- If OP interval exists, plugin adds `AniSkip Intro Start` and `AniSkip Intro End` chapters.
- At intro start, plugin shows an OSD hint for the first 3 seconds (`You can skip by pressing y-k` by default).
- Use `script-message subminer-aniskip-refresh` after changing media metadata/options to retry lookup.

## Lifecycle

- **File loaded**: If `auto_start=yes`, the plugin starts the overlay, then defers AniSkip lookup until after startup delay.
- **Auto-start pause gate**: If `auto_start_visible_overlay=yes` and `auto_start_pause_until_ready=yes`, launcher starts mpv paused and the plugin resumes playback after SubMiner reports tokenization-ready (with timeout fallback).
- **Duplicate auto-start events**: Repeated `file-loaded` hooks while overlay is already running are ignored for auto-start triggers (prevents duplicate start attempts).
- **MPV shutdown**: The plugin sends a stop command to gracefully shut down both the overlay and the texthooker server.
- **Texthooker**: Starts as a separate subprocess before the overlay to ensure the app lock is acquired first.

## Using with the `subminer` Wrapper

The `subminer` wrapper script handles mpv launch, socket setup, and overlay lifecycle automatically. You do not need the plugin if you always use the wrapper.

The plugin is useful when you:

- Launch mpv from other tools (file managers, media centers).
- Want on-demand overlay control without the wrapper.
- Use mpv's built-in file browser or playlist features.

You can install both — the plugin provides chord keybindings for convenience, while the wrapper handles the full lifecycle.
