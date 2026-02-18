# MPV Plugin

The SubMiner mpv plugin (`subminer.lua`) provides in-player keybindings to control the overlay without leaving mpv. It communicates with SubMiner by invoking the AppImage (or binary) with CLI flags.

## Installation

```bash
cp plugin/subminer.lua ~/.config/mpv/scripts/
cp plugin/subminer.conf ~/.config/mpv/script-opts/
# or: make install-plugin
```

mpv must have IPC enabled for SubMiner to connect:

```ini
# ~/.config/mpv/mpv.conf
input-ipc-server=/tmp/subminer-socket
```

## Keybindings

All keybindings use a `y` chord prefix — press `y`, then the second key:

| Chord | Action                   |
| ----- | ------------------------ |
| `y-y` | Open menu                |
| `y-s` | Start overlay            |
| `y-S` | Stop overlay             |
| `y-t` | Toggle visible overlay   |
| `y-i` | Toggle invisible overlay |
| `y-I` | Show invisible overlay   |
| `y-u` | Hide invisible overlay   |
| `y-o` | Open settings window     |
| `y-r` | Restart overlay          |
| `y-c` | Check status             |

## Menu

Press `y-y` to open an interactive menu in mpv's OSD:

```text
SubMiner:
1. Start overlay
2. Stop overlay
3. Toggle overlay
4. Toggle invisible overlay
5. Open options
6. Restart overlay
7. Check status
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
auto_start=no

# Show the visible overlay on auto-start.
auto_start_visible_overlay=no

# Invisible overlay startup: platform-default, visible, hidden.
# platform-default = hidden on Linux, visible on macOS/Windows.
auto_start_invisible_overlay=platform-default

# Show OSD messages for overlay status changes.
osd_messages=yes

# Logging level: debug, info, warn, error.
log_level=info
```

### Option Reference

| Option                         | Default                | Values                                     | Description                      |
| ------------------------------ | ---------------------- | ------------------------------------------ | -------------------------------- |
| `binary_path`                  | `""` (auto-detect)     | file path                                  | Path to SubMiner binary          |
| `socket_path`                  | `/tmp/subminer-socket` | file path                                  | MPV IPC socket path              |
| `texthooker_enabled`           | `yes`                  | `yes` / `no`                               | Enable texthooker server         |
| `texthooker_port`              | `5174`                 | 1–65535                                    | Texthooker server port           |
| `backend`                      | `auto`                 | `auto`, `hyprland`, `sway`, `x11`, `macos` | Window manager backend           |
| `auto_start`                   | `no`                   | `yes` / `no`                               | Auto-start overlay on file load  |
| `auto_start_visible_overlay`   | `no`                   | `yes` / `no`                               | Show visible layer on auto-start |
| `auto_start_invisible_overlay` | `platform-default`     | `platform-default`, `visible`, `hidden`    | Invisible layer on auto-start    |
| `osd_messages`                 | `yes`                  | `yes` / `no`                               | Show OSD status messages         |
| `log_level`                    | `info`                 | `debug`, `info`, `warn`, `error`           | Log verbosity                    |

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
script-message subminer-toggle-invisible
script-message subminer-show-invisible
script-message subminer-hide-invisible
script-message subminer-menu
script-message subminer-options
script-message subminer-restart
script-message subminer-status
```

The `subminer-start` message accepts overrides:

```
script-message subminer-start backend=hyprland socket=/custom/path texthooker=no log-level=debug
```

`log-level` here controls only logging verbosity passed to SubMiner.
`--debug` is a separate app/dev-mode flag in the main CLI and should not be used here for logging.

## Lifecycle

- **File loaded**: If `auto_start=yes`, the plugin starts the overlay and applies visibility preferences after a short delay.
- **MPV shutdown**: The plugin sends a stop command to gracefully shut down both the overlay and the texthooker server.
- **Texthooker**: Starts as a separate subprocess before the overlay to ensure the app lock is acquired first.

## Using with the `subminer` Wrapper

The `subminer` wrapper script handles mpv launch, socket setup, and overlay lifecycle automatically. You do not need the plugin if you always use the wrapper.

The plugin is useful when you:

- Launch mpv from other tools (file managers, media centers).
- Want on-demand overlay control without the wrapper.
- Use mpv's built-in file browser or playlist features.

You can install both — the plugin provides chord keybindings for convenience, while the wrapper handles the full lifecycle.
