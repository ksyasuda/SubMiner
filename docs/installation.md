# Requirements

### Linux

- **Wayland/X11 compositor** (one of the following):
  - Hyprland (uses `hyprctl`)
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
- yt-dlp (recommended for reliable YouTube playback/subtitles in mpv)
- bun (required to run the `subminer` wrapper script from source/local installs)

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

Note: the `subminer` wrapper uses a Bun shebang (`#!/usr/bin/env bun`), so `bun` must be installed and available on `PATH`.

### macOS Installation

If you download a release, use the **DMG** artifact. Open it and drag `SubMiner.app` into `/Applications`.
If needed, you can use the **ZIP** artifact as a fallback by unzipping and dragging `SubMiner.app` into `/Applications`.

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

The built app will be available in the `release` directory (`.dmg` and `.zip` on macOS).

If you are building locally without Apple signing credentials, use:

```bash
pnpm run build:mac:unsigned
```

You can launch `SubMiner.app` directly (double-click or `open -a SubMiner`).
Use `--start` when you want SubMiner to begin MPV IPC connection/reconnect behavior.
Use `--texthooker` when you only want the texthooker page (no overlay window).

**Accessibility Permission:**

After launching the app for the first time, grant accessibility permission:

1. Open **System Preferences** → **Security & Privacy** → **Privacy** tab
2. Select **Accessibility** from the left sidebar
3. Add SubMiner to the list

Without this permission, window tracking will not work and the overlay won't follow the MPV window.

<!-- ### Maintainer: macOS Release Signing/Notarization

The GitHub release workflow builds signed and notarized macOS artifacts on `macos-latest`.
Set these GitHub Actions secrets before creating a release tag:

- `CSC_LINK` (base64 `.p12` certificate or file URL for Developer ID Application cert)
- `CSC_KEY_PASSWORD` (password for the `.p12`)
- `APPLE_ID` (Apple ID email)
- `APPLE_APP_SPECIFIC_PASSWORD` (app-specific password for notarization)
- `APPLE_TEAM_ID` (Apple Developer Team ID) -->

### From Source

```bash
git clone https://github.com/sudacode/subminer.git
cd subminer
make build

# Install platform artifacts
# - Linux: wrapper + theme (+ AppImage if present)
# - macOS: wrapper + theme + SubMiner.app (from release/*.app or release/*.zip)
make install
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

Settings are stored in `~/.config/SubMiner/config.jsonc` (same as Linux).

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
| `y-t`   | Toggle visible overlay                |
| `y-i`   | Toggle invisible overlay              |
| `y-I`   | Show invisible overlay                |
| `y-u`   | Hide invisible overlay                |
| `y-o`   | Open Yomitan settings                 |
| `y-r`   | Restart overlay                       |
| `y-c`   | Check overlay status                  |

The menu provides options to start/stop/toggle the visible or invisible overlay layers and open settings. Type to filter or use arrow keys to navigate.

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

# Window manager backend: auto, hyprland, x11, macos
backend=auto

# Automatically start overlay when a file is loaded
auto_start=no

# Automatically show visible overlay when overlay starts
auto_start_visible_overlay=no

# Automatically show invisible overlay when overlay starts
# Values: platform-default, visible, hidden
# platform-default => hidden on Linux, visible on macOS/Windows
auto_start_invisible_overlay=platform-default

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

