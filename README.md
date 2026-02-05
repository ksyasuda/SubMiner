<div align="center">
  <img src="assets/subminer.png" width="169" alt="SubMiner logo">
  <h1>SubMiner</h1>
</div>

An all-in-one sentence mining overlay for MPV with AnkiConnect and dictionary (Yomitan) integration.

## Features

- Real-time subtitle display from MPV via IPC socket
- Yomitan integration for fast, on-screen lookups
- Japanese text tokenization using MeCab with smart word boundary detection
- Integrated texthooker-ui server for use with Yomitan
- Integrated websocket server (if [mpv_websocket](https://github.com/kuroahna/mpv_websocket) is not found) to send lines to the texthooker
- AnkiConnect integration for automatic card creation with media (audio/image)
- Secondary subtitle display with configurable display modes (hidden, visible, hover)

## Demo

<video src="https://github.com/user-attachments/assets/5f3705e5-af7c-4618-a433-c8fd30bf5db8">
  <source src="https://github.com/user-attachments/assets/5f3705e5-af7c-4618-a433-c8fd30bf5db8" type="video/mp4">
  Your browser does not support the video tag.
</video>

## Requirements

### Linux
- **Wayland/X11 compositor** (one of the following):
  - Hyprland (uses `hyprctl`)
  - Sway (uses `swaymsg`)
  - X11 (uses `xdotool` and `xwininfo`)
- mpv (with IPC socket support)
- mecab and mecab-ipadic (Japanese morphological analyzer)
- fuse2 (for AppImage support)

### macOS
- macOS 10.13 or later
- mpv (with IPC socket support)
- mecab and mecab-ipadic (Japanese morphological analyzer) - optional
- **Accessibility permission** required for window tracking (see [macOS Installation](#macos-installation))

**Optional:**

- fzf (terminal-based video picker, default)
- rofi (GUI-based video picker)
- chafa (thumbnail previews in fzf)
- ffmpegthumbnailer (generate video thumbnails)

## Installation

### From AppImage (Recommended)

Download the latest AppImage from GitHub Releases:

```bash
# Download and install AppImage
wget https://github.com/sudacode/subminer/releases/download/v1.0.0/subminer-1.0.0.AppImage -O ~/.local/bin/subminer.AppImage
chmod +x ~/.local/bin/subminer.AppImage

# Download subminer wrapper script
wget https://github.com/sudacode/subminer/releases/download/v1.0.0/subminer -O ~/.local/bin/subminer
chmod +x ~/.local/bin/subminer
```

### macOS Installation

If you download a release, use the **ZIP** artifact. Unzip it and drag `SubMiner.app` into `/Applications`.

Install dependencies using Homebrew:

```bash
brew install mpv mecab mecab-ipadic
```

Build from source:

```bash
git clone https://github.com/sudacode/subminer.git
cd subminer
pnpm install
cd vendor/texthooker-ui && pnpm install && pnpm build && cd ../..
pnpm run build:mac
```

The built app will be available in the `release` directory (ZIP on macOS).

You can launch `SubMiner.app` directly (double-click or `open -a SubMiner`). The app no longer requires a `--start` argument on macOS.

**Accessibility Permission:**

After launching the app for the first time, grant accessibility permission:
1. Open **System Preferences** → **Security & Privacy** → **Privacy** tab
2. Select **Accessibility** from the left sidebar
3. Add SubMiner to the list

Without this permission, window tracking will not work and the overlay won't follow the MPV window.

### From Source (Linux/Development)

```bash
git clone https://github.com/sudacode/subminer.git
cd subminer
pnpm install
cd vendor/texthooker-ui && pnpm install && pnpm build && cd ../..
pnpm run build:appimage

# Install wrapper script + rofi theme (and AppImage if present)
make install-linux
```

<!-- ### Arch Linux -->

<!-- ```bash -->
<!-- # Using the PKGBUILD -->
<!-- makepkg -si -->
<!-- ``` -->

### macOS Usage Notes

**Launching MPV with IPC:**

```bash
mpv --input-ipc-server=/tmp/subminer-socket video.mkv
```

**Config Location:**

Settings are stored in `~/.config/subminer/config.json` (same as Linux).

**MeCab Installation Paths:**

Common Homebrew install paths:
- Apple Silicon (M1/M2): `/opt/homebrew/bin/mecab`
- Intel: `/usr/local/bin/mecab`

Ensure that `mecab` is available on your PATH when launching subminer (for example, by starting it from a terminal where `which mecab` works), otherwise MeCab may not be detected.

**Fullscreen Mode:**

The overlay should appear correctly in fullscreen. If you encounter issues, check that macOS accessibility permissions are granted (see [macOS Installation](#macos-installation)).

**mpv Plugin Binary Path (macOS):**

Set `binary_path` to your app binary, for example:

```ini
binary_path=/Applications/SubMiner.app/Contents/MacOS/subminer
```

### MPV Plugin (Optional)

The Lua plugin allows you to control the overlay directly from mpv using keybindings:

> [!IMPORTANT]
> `mpv` must be launched with `--input-ipc-server=/tmp/subminer-socket` to allow communication with the application

```bash
# Copy plugin files to mpv config
cp plugin/subminer.lua ~/.config/mpv/scripts/
cp plugin/subminer.conf ~/.config/mpv/script-opts/
```

#### Plugin Keybindings

All keybindings use chord sequences starting with `y`:

| Keybind | Action                                |
| ------- | ------------------------------------- |
| `y-y`   | Open SubMiner menu (fuzzy-searchable) |
| `y-s`   | Start overlay                         |
| `y-S`   | Stop overlay                          |
| `y-t`   | Toggle overlay                        |
| `y-o`   | Open Yomitan settings                 |
| `y-r`   | Restart overlay                       |
| `y-c`   | Check overlay status                  |

The menu provides options to start/stop/toggle the overlay and open settings. Type to filter or use arrow keys to navigate.

#### Plugin Configuration

Edit `~/.config/mpv/script-opts/subminer.conf`:

```ini
# Path to SubMiner binary (leave empty for auto-detection)
binary_path=

# Path to mpv IPC socket (must match input-ipc-server in mpv.conf)
socket_path=/tmp/subminer-socket

# Enable texthooker WebSocket server
texthooker_enabled=yes

# Texthooker WebSocket port
texthooker_port=5174

# Window manager backend: auto, hyprland, sway, x11, macos
backend=auto

# Automatically start overlay when a file is loaded
auto_start=no

# Automatically show overlay (hide mpv subtitles) when overlay starts
auto_start_overlay=yes

# Show OSD messages for overlay status
osd_messages=yes
```

The plugin auto-detects the binary location, searching:

- `/Applications/SubMiner.app/Contents/MacOS/subminer`
- `~/Applications/SubMiner.app/Contents/MacOS/subminer`
- `C:\Program Files\subminer\subminer.exe`
- `C:\Program Files (x86)\subminer\subminer.exe`
- `C:\subminer\subminer.exe`
- `~/.local/bin/subminer.AppImage`
- `/opt/subminer/subminer.AppImage`
- `/usr/local/bin/subminer`
- `/usr/bin/subminer`

**Windows Notes:**

Set the binary and socket path like this:

```ini
binary_path=C:\\Program Files\\subminer\\subminer.exe
socket_path=\\\\.\\pipe\\subminer-socket
```

Launch mpv with:

```bash
mpv --input-ipc-server=\\\\.\\pipe\\subminer-socket video.mkv
```

## SubMiner Script vs MPV Plugin

There are two ways to use SubMiner:

| Approach            | Best For                                                                                                                                                                              |
| ------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **subminer script** | All-in-one solution. Handles video selection, launches MPV with the correct socket, starts the overlay automatically, and cleans up on exit.                                          |
| **MPV plugin**      | When you launch MPV yourself or from other tools. Provides in-MPV chord keybindings (e.g. `y-y` for menu) to control the overlay. Requires `--input-ipc-server=/tmp/subminer-socket`. |

You can use both together—install the plugin for on-demand control, but use `subminer` when you want the streamlined workflow.

## Usage

```bash
# Browse and play videos
subminer                          # Current directory (uses fzf)
subminer -R                       # Use rofi instead of fzf
subminer -d ~/Videos              # Specific directory
subminer -r -d ~/Anime            # Recursive search
subminer video.mkv                # Play specific file

# Options
subminer -T video.mkv             # Disable texthooker server
subminer -b x11 video.mkv         # Force X11 backend
subminer -p gpu-hq video.mkv      # Use mpv profile

# Direct AppImage control
subminer.AppImage --start --texthooker   # Start overlay with texthooker
subminer.AppImage --stop                  # Stop overlay
subminer.AppImage --toggle                # Toggle visibility
subminer.AppImage --settings              # Open Yomitan settings
subminer.AppImage --help                  # Show all options
```

## Keybindings

### Global Shortcuts

| Keybind       | Action                    |
| ------------- | ------------------------- |
| `Alt+Shift+O` | Toggle overlay visibility |
| `Alt+Shift+Y` | Open Yomitan settings     |

### Overlay Controls (Configurable)

| Input                | Action                                             |
| -------------------- | -------------------------------------------------- |
| `Space`              | Toggle MPV pause                                   |
| `Shift+H`            | Jump to previous subtitle                          |
| `Shift+L`            | Jump to next subtitle                              |
| `Ctrl+Shift+H`       | Replay current subtitle (play to end, then pause)  |
| `Ctrl+Shift+L`       | Play next subtitle (jump, play to end, then pause) |
| `Right-click`        | Toggle MPV pause (outside subtitle area)           |
| `Right-click + drag` | Move subtitle position (on subtitle)               |

These keybindings only work when the overlay window has focus. See [Configuration](#configuration) for customization.

### Overlay Chord Shortcuts

| Chord | Action |
| ----- | ------ |
| `y` → `j` | Open Jimaku subtitle menu |

## How It Works

1. MPV runs with an IPC socket at `/tmp/subminer-socket`
2. The overlay connects and subscribes to subtitle changes
3. Subtitles are tokenized with MeCab and merged into natural word boundaries
4. Words are displayed as clickable spans
5. Clicking a word triggers Yomitan popup for dictionary lookup
6. Texthooker server runs at `http://127.0.0.1:5174` for external tools

## Configuration

Settings are stored in `~/.config/subminer/config.json`

### Configuration File

See `config.example.jsonc` for a comprehensive example configuration file with all available options, default values, and detailed comments. Only include the options you want to customize in your config file.

### Configuration Options Overview

The configuration file includes several main sections:

- **Texthooker** - Control browser opening behavior
- **WebSocket** - Built-in subtitle broadcasting server
- **AnkiConnect** - Automatic Anki card creation with media
- **Shortcuts** - Overlay keyboard shortcuts
- **Keybindings** - MPV command shortcuts
- **Subtitle Style** - Appearance customization
- **Secondary Subtitles** - Dual subtitle track support
- **Subtitle Position** - Overlay vertical positioning

### Texthooker

Control whether the browser opens automatically when texthooker starts:

See `config.example.jsonc` for detailed configuration options.

```json
{
  "texthooker": {
    "openBrowser": true
  }
}
```

### Jimaku

Configure Jimaku API access and defaults:

```json
{
  "jimaku": {
    "apiKey": "YOUR_API_KEY",
    "apiKeyCommand": "cat ~/.jimaku_key",
    "apiBaseUrl": "https://jimaku.cc",
    "languagePreference": "ja",
    "maxEntryResults": 10
  }
}
```

Jimaku is rate limited; if you hit a limit, SubMiner will surface the retry delay from the API response.

Set `openBrowser` to `false` to only print the URL without opening a browser.

### WebSocket Server

The overlay includes a built-in WebSocket server that broadcasts subtitle text to connected clients (such as texthooker-ui) for external processing.

By default, the server uses "auto" mode: it starts automatically unless [mpv_websocket](https://github.com/kuroahna/mpv_websocket) is detected at `~/.config/mpv/mpv_websocket`. If you have mpv_websocket installed, the built-in server is skipped to avoid conflicts.

See `config.example.jsonc` for detailed configuration options.

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

### AnkiConnect

Enable automatic Anki card creation and updates with media generation:

See `config.example.jsonc` for detailed configuration options with all available parameters.

```json
{
  "ankiConnect": {
    "enabled": true,
    "url": "http://127.0.0.1:8765",
    "pollingRate": 3000,
    "deck": "Learning::Japanese",
    "audioField": "ExpressionAudio",
    "imageField": "Picture",
    "sentenceField": "Sentence"
    // ... many more options available in config.example.jsonc
  }
}
```

**Requirements:** [AnkiConnect](https://github.com/FooSoft/anki-connect) plugin must be installed and running in Anki. ffmpeg must be installed for media generation.

| Option                      | Values                                  | Description                                                                                                                 |
| --------------------------- | --------------------------------------- | --------------------------------------------------------------------------------------------------------------------------- |
| `enabled`                   | `true`, `false`                         | Enable AnkiConnect integration (default: `false`)                                                                           |
| `url`                       | string (URL)                            | AnkiConnect API URL (default: `http://127.0.0.1:8765`)                                                                      |
| `pollingRate`               | number (ms)                             | How often to check for new cards (default: `3000`)                                                                          |
| `deck`                      | string                                  | Anki deck to monitor for new cards                                                                                          |
| `audioField`                | string                                  | Card field for audio files (default: `ExpressionAudio`)                                                                     |
| `imageField`                | string                                  | Card field for images (default: `Picture`)                                                                                  |
| `sentenceField`             | string                                  | Card field for sentences (default: `Sentence`)                                                                              |
| `generateAudio`             | `true`, `false`                         | Generate audio clips from video (default: `true`)                                                                           |
| `generateImage`             | `true`, `false`                         | Generate image/animation screenshots (default: `true`)                                                                      |
| `imageType`                 | `"static"`, `"avif"`                    | Image type: static screenshot or animated AVIF (default: `"static"`)                                                        |
| `imageFormat`               | `"jpg"`, `"png"`, `"webp"`              | Image format (default: `"jpg"`)                                                                                             |
| `imageQuality`              | number (1-100)                          | Image quality for JPG/WebP; PNG ignores this (default: `92`)                                                                |
| `imageMaxWidth`             | number (px)                             | Max width for images; preserves aspect ratio (default: `1280`)                                                              |
| `imageMaxHeight`            | number (px)                             | Max height for images; preserves aspect ratio (default: `720`)                                                              |
| `animatedFps`               | number (1-60)                           | FPS for animated AVIF (default: `10`)                                                                                       |
| `animatedMaxWidth`          | number (px)                             | Max width for animated AVIF (default: `640`)                                                                                |
| `animatedMaxHeight`         | number (px)                             | Max height for animated AVIF; preserves aspect ratio (default: `null`)                                                      |
| `animatedCrf`               | number (0-63)                           | CRF quality for AVIF; lower = higher quality (default: `35`)                                                                |
| `audioPadding`              | number (seconds)                        | Padding around audio clip timing (default: `0.5`)                                                                           |
| `fallbackDuration`          | number (seconds)                        | Default duration if timing unavailable (default: `3.0`)                                                                     |
| `overwriteAudio`            | `true`, `false`                         | Replace existing audio on updates; when `false`, new audio is appended/prepended per `mediaInsertMode` (default: `true`)    |
| `overwriteImage`            | `true`, `false`                         | Replace existing images on updates; when `false`, new images are appended/prepended per `mediaInsertMode` (default: `true`) |
| `mediaInsertMode`           | `"append"`, `"prepend"`                 | Where to insert new media when overwrite is off (default: `"append"`)                                                       |
| `miscInfoField`             | string                                  | Card field for metadata (optional)                                                                                          |
| `miscInfoPattern`           | string                                  | Format pattern for metadata: `%f`=filename, `%F`=filename+ext, `%t`=time                                                    |
| `highlightWord`             | `true`, `false`                         | Highlight the word in sentence context (default: `true`)                                                                    |
| `notificationType`          | `"osd"`, `"system"`, `"both"`, `"none"` | Notification type on card update (default: `"osd"`)                                                                         |
| `autoUpdateNewCards`        | `true`, `false`                         | Automatically update cards on creation (default: `true`)                                                                    |
| `maxMediaDuration`          | number (seconds)                        | Max duration for generated media from multi-line copy (default: `30`, `0` to disable)                                       |
| `sentenceCardModel`         | string                                  | Anki note type for sentence mining cards (optional)                                                                         |
| `sentenceCardSentenceField` | string                                  | Field name for sentence in sentence cards (default: `Sentence`)                                                             |
| `sentenceCardAudioField`    | string                                  | Field name for audio in sentence cards (default: `SentenceAudio`)                                                           |
| `isLapis`                   | `true`, `false`                         | Enable Lapis note format compatibility (default: `false`)                                                                   |

**Image Quality Notes:**

- `imageQuality` affects JPG and WebP only; PNG is lossless and ignores this setting
- JPG quality is mapped to FFmpeg's scale (2-31, lower = better)
- WebP quality uses FFmpeg's native 0-100 scale

**Requirements:** [AnkiConnect](https://github.com/FooSoft/anki-connect) plugin must be installed and running in Anki. ffmpeg must be installed for media generation.

**See `config.example.jsonc`** for the complete list of AnkiConnect configuration options.

**Manual Card Update:**

When `autoUpdateNewCards` is set to `false`, new cards are detected but not automatically updated. Instead, you can manually update cards using keyboard shortcuts:

| Shortcut       | Action                                                                                                       |
| -------------- | ------------------------------------------------------------------------------------------------------------ |
| `Ctrl+C`       | Copy the current subtitle line to clipboard (preserves line breaks)                                          |
| `Ctrl+Shift+C` | Enter multi-copy mode. Press `1-9` to copy that many recent lines, or `Esc` to cancel. Timeout: 3 seconds    |
| `Ctrl+V`       | Update the last added Anki card using subtitles from clipboard                                               |
| `Ctrl+S`       | Create a sentence card from the current subtitle line                                                        |
| `Ctrl+Shift+S` | Enter multi-mine mode. Press `1-9` to create a sentence card from that many recent lines, or `Esc` to cancel |
| `Ctrl+Shift+V` | Cycle secondary subtitle display mode (hidden → visible → hover)                                             |

To copy multiple lines (current + previous):

1. Press `Ctrl+Shift+C`
2. Press a number key (`1-9`) within 3 seconds
3. The specified number of most recent subtitle lines are copied
4. Press `Ctrl+V` to update the last added card with the copied lines

These shortcuts are only active when the overlay window is visible. They are automatically disabled when the overlay is hidden to avoid interfering with normal system clipboard operations.

### Shortcuts Configuration

Customize or disable the overlay keyboard shortcuts:

See `config.example.jsonc` for detailed configuration options.

```json
{
  "shortcuts": {
    "copySubtitle": "CommandOrControl+C",
    "copySubtitleMultiple": "CommandOrControl+Shift+C",
    "updateLastCardFromClipboard": "CommandOrControl+V",
    "mineSentence": "CommandOrControl+S",
    "mineSentenceMultiple": "CommandOrControl+Shift+S",
    "multiCopyTimeoutMs": 3000
  }
}
```

| Option                        | Values           | Description                                                                                    |
| ----------------------------- | ---------------- | ---------------------------------------------------------------------------------------------- |
| `copySubtitle`                | string \| `null` | Accelerator for copying current subtitle (default: `"CommandOrControl+C"`)                     |
| `copySubtitleMultiple`        | string \| `null` | Accelerator for multi-copy mode (default: `"CommandOrControl+Shift+C"`)                        |
| `updateLastCardFromClipboard` | string \| `null` | Accelerator for updating card from clipboard (default: `"CommandOrControl+V"`)                 |
| `mineSentence`                | string \| `null` | Accelerator for creating sentence card from current subtitle (default: `"CommandOrControl+S"`) |
| `mineSentenceMultiple`        | string \| `null` | Accelerator for multi-mine sentence card mode (default: `"CommandOrControl+Shift+S"`)          |
| `multiCopyTimeoutMs`          | number           | Timeout in ms for multi-copy/mine digit input (default: `3000`)                                |
| `toggleSecondarySub`          | string \| `null` | Accelerator for cycling secondary subtitle mode (default: `"CommandOrControl+Shift+V"`)        |

**See `config.example.jsonc`** for the complete list of shortcut configuration options. |

Set any shortcut to `null` to disable it.

### Keybindings

Add a `keybindings` array to configure keyboard shortcuts that send commands to mpv:

See `config.example.jsonc` for detailed configuration options and more examples.

**Default keybindings:**

| Key               | Command                    | Description                           |
| ----------------- | -------------------------- | ------------------------------------- |
| `Space`           | `["cycle", "pause"]`       | Toggle pause                          |
| `Shift+KeyH`      | `["sub-seek", -1]`         | Jump to previous subtitle             |
| `Shift+KeyL`      | `["sub-seek", 1]`          | Jump to next subtitle                 |
| `Ctrl+Shift+KeyH` | `["__replay-subtitle"]`    | Replay current subtitle, pause at end |
| `Ctrl+Shift+KeyL` | `["__play-next-subtitle"]` | Play next subtitle, pause at end      |

**Custom keybindings example:**

```json
{
  "keybindings": [
    { "key": "ArrowRight", "command": ["seek", 5] },
    { "key": "ArrowLeft", "command": ["seek", -5] },
    { "key": "Shift+ArrowRight", "command": ["seek", 30] },
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

**Special commands:** Commands prefixed with `__` are handled internally by the overlay rather than sent to mpv. `__replay-subtitle` replays the current subtitle and pauses at its end. `__play-next-subtitle` seeks to the next subtitle, plays it, and pauses at its end.

**Supported commands:** Any valid mpv JSON IPC command array (`["cycle", "pause"]`, `["seek", 5]`, `["script-binding", "..."]`, etc.)

**See `config.example.jsonc`** for more keybinding examples and configuration options.

### Subtitle Style

Customize the appearance of primary and secondary subtitles:

See `config.example.jsonc` for detailed configuration options.

```json
{
  "subtitleStyle": {
    "fontFamily": "Noto Sans CJK JP Regular, Noto Sans CJK JP, Arial Unicode MS, Arial, sans-serif",
    "fontSize": 35,
    "fontColor": "#cad3f5",
    "fontWeight": "normal",
    "fontStyle": "normal",
    "backgroundColor": "rgba(54, 58, 79, 0.5)",
    "secondary": {
      "fontSize": 24,
      "fontColor": "#cad3f5",
      "backgroundColor": "transparent"
    }
  }
}
```

| Option            | Values      | Description                                                                   |
| ----------------- | ----------- | ----------------------------------------------------------------------------- |
| `fontFamily`      | string      | CSS font-family value (default: `"Noto Sans CJK JP Regular, ..."`)            |
| `fontSize`        | number (px) | Font size in pixels (default: `35`)                                           |
| `fontColor`       | string      | Any CSS color value (default: `"#cad3f5"`)                                    |
| `fontWeight`      | string      | CSS font-weight, e.g. `"bold"`, `"normal"`, `"600"` (default: `"normal"`)     |
| `fontStyle`       | string      | `"normal"` or `"italic"` (default: `"normal"`)                                |
| `backgroundColor` | string      | Any CSS color, including `"transparent"` (default: `"rgba(54, 58, 79, 0.5)"`) |
| `secondary`       | object      | Override any of the above for secondary subtitles (optional)                  |

Secondary subtitle defaults: `fontSize: 24`, `backgroundColor: "transparent"`. Any property not set in `secondary` falls back to the CSS defaults.

**See `config.example.jsonc`** for the complete list of subtitle style configuration options.

### Secondary Subtitles

Display a second subtitle track (e.g., English alongside Japanese) in the overlay:

See `config.example.jsonc` for detailed configuration options.

```json
{
  "secondarySub": {
    "secondarySubLanguages": ["eng", "en"],
    "autoLoadSecondarySub": true,
    "defaultMode": "hover"
  }
}
```

| Option                  | Values                             | Description                                            |
| ----------------------- | ---------------------------------- | ------------------------------------------------------ |
| `secondarySubLanguages` | string[]                           | Language codes to auto-load (e.g., `["eng", "en"]`)    |
| `autoLoadSecondarySub`  | `true`, `false`                    | Auto-detect and load matching secondary subtitle track |
| `defaultMode`           | `"hidden"`, `"visible"`, `"hover"` | Initial display mode (default: `"hover"`)              |

**Display modes:**

- **hidden** — Secondary subtitles not shown
- **visible** — Always visible at top of overlay
- **hover** — Only visible when hovering over the subtitle area (default)

**See `config.example.jsonc`** for additional secondary subtitle configuration options.

## Environment Variables

| Variable                 | Description                                    |
| ------------------------ | ---------------------------------------------- |
| `SUBMINER_APPIMAGE_PATH` | Override AppImage location for subminer script |

## License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.

### Third-Party Components

This project includes the following third-party components:

- **[Yomitan](https://github.com/yomidevs/yomitan)** - GPL-3.0
- **[texthooker-ui](https://github.com/Renji-XD/texthooker-ui)** - MIT

### Acknowledgments

This project was inspired by **[GameSentenceMiner](https://github.com/bpwhelan/GameSentenceMiner)**'s subtitle overlay and Yomitan integration
