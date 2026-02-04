<div align="center">
  <img src="assets/mpv-yomitan.png" width="169" height="169" alt="mpv-yomitan logo">
</div>

# mpv-yomitan
An Electron-based subtitle overlay for MPV with Yomitan lookup support for Japanese language learning.

## Features

- Real-time subtitle display from MPV via IPC socket
- Yomitan integration for fast, on-screen lookups
- Japanese text tokenization using MeCab with smart word boundary detection
- Integrated texthooker-ui server for use with Yomitan
- Integrated websocket server (if [mpv_websocket](https://github.com/kuroahna/mpv_websocket) is not found) to send lines to the texthooker

## Demo

<video src="https://github.com/user-attachments/assets/5f3705e5-af7c-4618-a433-c8fd30bf5db8">
  <source src="https://github.com/user-attachments/assets/5f3705e5-af7c-4618-a433-c8fd30bf5db8" type="video/mp4">
  Your browser does not support the video tag.
</video>

## Requirements

- **Wayland/X11 compositor** (one of the following):
  - Hyprland (uses `hyprctl`)
  - Sway (uses `swaymsg`)
  - X11 (uses `xdotool` and `xwininfo`)
- mpv (with IPC socket support)
- mecab and mecab-ipadic (Japanese morphological analyzer)
- fuse2 (for AppImage support)

**Optional:**

- fzf (terminal-based video picker, default)
- rofi (GUI-based video picker)
- chafa (thumbnail previews in fzf)
- ffmpegthumbnailer (generate video thumbnails)

## Installation

### From AppImage (Recommended)

Download the latest AppImage from [Releases](https://github.com/ksyasuda/mpv-yomitan/releases):

```bash
# Download and install AppImage
wget https://github.com/ksyasuda/mpv-yomitan/releases/download/v1.0.0/mpv-yomitan-1.0.0.AppImage -O ~/.local/bin/mpv-yomitan.AppImage
chmod +x ~/.local/bin/mpv-yomitan.AppImage

# Download ympv wrapper script
wget https://github.com/ksyasuda/mpv-yomitan/releases/download/v1.0.0/ympv -O ~/.local/bin/ympv
chmod +x ~/.local/bin/ympv
```

### From Source (Development)

```bash
git clone https://github.com/ksyasuda/mpv-yomitan.git
cd mpv-yomitan
pnpm install
cd vendor/texthooker-ui && pnpm install && pnpm build && cd ../..
pnpm run build:appimage

# Copy to ~/.local/bin
cp dist/mpv-yomitan-*.AppImage ~/.local/bin/mpv-yomitan.AppImage
cp ympv ~/.local/bin/
chmod +x ~/.local/bin/mpv-yomitan.AppImage ~/.local/bin/ympv
```

<!-- ### Arch Linux -->

<!-- ```bash -->
<!-- # Using the PKGBUILD -->
<!-- makepkg -si -->
<!-- ``` -->

### MPV Plugin (Optional)

The Lua plugin allows you to control the overlay directly from mpv using keybindings:

> [!IMPORTANT]
> `mpv` must be launched with `--input-ipc-server=/tmp/mpv-yomitan-socket` to allow communication with the application

```bash
# Copy plugin files to mpv config
cp plugin/mpv-yomitan.lua ~/.config/mpv/scripts/
cp plugin/mpv-yomitan.conf ~/.config/mpv/script-opts/
```

#### Plugin Keybindings

All keybindings use chord sequences starting with `y`:

| Keybind | Action                                   |
| ------- | ---------------------------------------- |
| `y-y`   | Open mpv-yomitan menu (fuzzy-searchable) |
| `y-s`   | Start overlay                            |
| `y-S`   | Stop overlay                             |
| `y-t`   | Toggle overlay                           |
| `y-o`   | Open Yomitan settings                    |
| `y-r`   | Restart overlay                          |
| `y-c`   | Check overlay status                     |

The menu provides options to start/stop/toggle the overlay and open settings. Type to filter or use arrow keys to navigate.

#### Plugin Configuration

Edit `~/.config/mpv/script-opts/mpv-yomitan.conf`:

```ini
# Path to mpv-yomitan binary (leave empty for auto-detection)
binary_path=

# Path to mpv IPC socket (must match input-ipc-server in mpv.conf)
socket_path=/tmp/mpv-yomitan-socket

# Enable texthooker WebSocket server
texthooker_enabled=yes

# Texthooker WebSocket port
texthooker_port=5174

# Window manager backend: auto, hyprland, sway, x11
backend=auto

# Automatically start overlay when a file is loaded
auto_start=no

# Automatically show overlay (hide mpv subtitles) when overlay starts
auto_start_overlay=yes

# Show OSD messages for overlay status
osd_messages=yes
```

The plugin auto-detects the binary location, searching:

- `~/.local/bin/mpv-yomitan.AppImage`
- `/opt/mpv-yomitan/mpv-yomitan`
- `/usr/local/bin/mpv-yomitan`
- `/usr/bin/mpv-yomitan`

## ympv Script vs MPV Plugin

There are two ways to use mpv-yomitan:

| Approach        | Best For                                                                                                                                                             |
| --------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **ympv script** | All-in-one solution. Handles video selection, launches MPV with the correct socket, starts the overlay automatically, and cleans up on exit.                         |
| **MPV plugin**  | When you launch MPV yourself or from other tools. Provides in-MPV chord keybindings (e.g. `y-y` for menu) to control the overlay. Requires `--input-ipc-server=/tmp/mpv-yomitan-socket`. |

You can use both together—install the plugin for on-demand control, but use `ympv` when you want the streamlined workflow.

## Usage

```bash
# Browse and play videos
ympv                          # Current directory (uses fzf)
ympv -R                       # Use rofi instead of fzf
ympv -d ~/Videos              # Specific directory
ympv -r -d ~/Anime            # Recursive search
ympv video.mkv                # Play specific file

# Options
ympv -T video.mkv             # Disable texthooker server
ympv -b x11 video.mkv         # Force X11 backend
ympv -p gpu-hq video.mkv      # Use mpv profile

# Direct AppImage control
mpv-yomitan.AppImage --start --texthooker   # Start overlay with texthooker
mpv-yomitan.AppImage --stop                  # Stop overlay
mpv-yomitan.AppImage --toggle                # Toggle visibility
mpv-yomitan.AppImage --settings              # Open Yomitan settings
mpv-yomitan.AppImage --help                  # Show all options
```

## Keybindings

### Global Shortcuts

| Keybind       | Action                    |
| ------------- | ------------------------- |
| `Alt+Shift+O` | Toggle overlay visibility |
| `Alt+Shift+Y` | Open Yomitan settings     |

### Overlay Controls (Configurable)

| Input                | Action                                   |
| -------------------- | ---------------------------------------- |
| `Space`              | Toggle MPV pause                         |
| `Right-click`        | Toggle MPV pause (outside subtitle area) |
| `Right-click + drag` | Move subtitle position (on subtitle)     |

These keybindings only work when the overlay window has focus. See [Configuration](#configuration) for customization.

## How It Works

1. MPV runs with an IPC socket at `/tmp/mpv-yomitan-socket`
2. The overlay connects and subscribes to subtitle changes
3. Subtitles are tokenized with MeCab and merged into natural word boundaries
4. Words are displayed as clickable spans
5. Clicking a word triggers Yomitan popup for dictionary lookup
6. Texthooker server runs at `http://127.0.0.1:5174` for external tools

## Configuration

Settings are stored in `~/.config/mpv-yomitan-overlay/config.json`

### Texthooker

Control whether the browser opens automatically when texthooker starts:

```json
{
  "texthooker": {
    "openBrowser": true
  }
}
```

Set `openBrowser` to `false` to only print the URL without opening a browser.

### WebSocket Server

The overlay includes a built-in WebSocket server that broadcasts subtitle text to connected clients (such as texthooker-ui) for external processing.

By default, the server uses "auto" mode: it starts automatically unless [mpv_websocket](https://github.com/kuroahna/mpv_websocket) is detected at `~/.config/mpv/mpv_websocket`. If you have mpv_websocket installed, the built-in server is skipped to avoid conflicts.

```json
{
  "websocket": {
    "enabled": "auto",
    "port": 6677
  }
}
```

| Option    | Values                    | Description                                              |
| --------- | ------------------------- | -------------------------------------------------------- |
| `enabled` | `true`, `false`, `"auto"` | `"auto"` (default) disables if mpv_websocket is detected |
| `port`    | number                    | WebSocket server port (default: 6677)                    |

### Keybindings

Add a `keybindings` array to configure keyboard shortcuts that send commands to mpv:

```json
{
  "subtitlePosition": { "yPercent": 10 },
  "keybindings": [
    { "key": "Space", "command": ["cycle", "pause"] },
    { "key": "ArrowRight", "command": ["seek", 5] },
    { "key": "ArrowLeft", "command": ["seek", -5] },
    { "key": "Shift+ArrowRight", "command": ["seek", 30] },
    { "key": "KeyJ", "command": ["sub-seek", -1] },
    { "key": "KeyL", "command": ["sub-seek", 1] },
    { "key": "KeyR", "command": ["script-binding", "immersive/auto-replay"] },
    { "key": "KeyA", "command": ["script-message", "ankiconnect-add-note"] }
  ]
}
```

**Key format:** Use `KeyboardEvent.code` values (`Space`, `ArrowRight`, `KeyR`, etc.) with optional modifiers (`Ctrl+`, `Alt+`, `Shift+`, `Meta+`).

**Disable a default binding:** Set command to `null`:

```json
{ "key": "Space", "command": null }
```

**Supported commands:** Any valid mpv JSON IPC command array (`["cycle", "pause"]`, `["seek", 5]`, `["script-binding", "..."]`, etc.)

## Environment Variables

| Variable                    | Description                                |
| --------------------------- | ------------------------------------------ |
| `MPV_YOMITAN_APPIMAGE_PATH` | Override AppImage location for ympv script |

## License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.

### Third-Party Components

This project includes the following third-party components:

- **[Yomitan](https://github.com/yomidevs/yomitan)** - GPL-3.0
- **[texthooker-ui](https://github.com/Renji-XD/texthooker-ui)** - MIT

### Acknowledgments

This project was inspired by **[GameSentenceMiner](https://github.com/bpwhelan/GameSentenceMiner)**'s subtitle overlay and Yomitan integration
