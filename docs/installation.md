# Installation

## Requirements

### System Dependencies

| Dependency           | Required   | Notes                                                    |
| -------------------- | ---------- | -------------------------------------------------------- |
| Bun                  | Yes        | Required for `subminer` wrapper and source workflows     |
| mpv                  | Yes        | Must support IPC sockets (`--input-ipc-server`)          |
| ffmpeg               | For media  | Audio extraction and screenshot generation               |
| MeCab + mecab-ipadic | No         | Optional fallback tokenizer for Japanese                 |
| fuse2                | Linux only | Required for AppImage                                    |
| yt-dlp               | No         | Recommended for YouTube playback and subtitle extraction |

### Platform-Specific

**Linux** — one of the following compositors:

- Hyprland (uses `hyprctl`)
- Sway (uses `swaymsg`)
- X11 (uses `xdotool` and `xwininfo`)

**macOS** — macOS 10.13 or later. Accessibility permission required for window tracking.

### Optional Tools

| Tool              | Purpose                                                       |
| ----------------- | ------------------------------------------------------------- |
| fzf               | Terminal-based video picker (default)                         |
| rofi              | GUI-based video picker                                        |
| chafa             | Thumbnail previews in fzf                                     |
| ffmpegthumbnailer | Generate video thumbnails for picker                          |
| guessit           | Better AniSkip title/season/episode parsing for file playback |
| alass             | Subtitle sync engine (preferred)                              |
| ffsubsync         | Subtitle sync engine (fallback)                               |

## Linux

### AppImage (Recommended)

Download the latest AppImage from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest):

```bash
# Download and install AppImage
wget https://github.com/ksyasuda/SubMiner/releases/download/v0.1.0/SubMiner-0.1.0.AppImage -O ~/.local/bin/SubMiner.AppImage
chmod +x ~/.local/bin/SubMiner.AppImage

# Download subminer wrapper script
wget https://github.com/ksyasuda/SubMiner/releases/download/v0.1.0/subminer -O ~/.local/bin/subminer
chmod +x ~/.local/bin/subminer
```

The `subminer` wrapper uses a Bun shebang (`#!/usr/bin/env bun`), so [Bun](https://bun.sh) must be installed and available on `PATH`.

### From Source

```bash
git clone https://github.com/ksyasuda/SubMiner.git
cd SubMiner
make build
make build-launcher

# Install platform artifacts (wrapper + theme + AppImage)
make install
```

`make build-launcher` generates the wrapper at `dist/launcher/subminer`. The checked-in launcher source remains `launcher/*.ts`.
Do not use a repo-root `./subminer` artifact when building from source; workflow checks enforce `dist/launcher/subminer` as the only generated path.

## macOS

### DMG (Recommended)

Download the **DMG** artifact from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest). Open it and drag `SubMiner.app` into `/Applications`.

A **ZIP** artifact is also available as a fallback — unzip and drag `SubMiner.app` into `/Applications`.

Install dependencies using Homebrew:

```bash
brew install mpv mecab mecab-ipadic
```

### From Source (macOS)

```bash
git clone https://github.com/ksyasuda/SubMiner.git
cd SubMiner
bun install
cd vendor/texthooker-ui && bun install --frozen-lockfile && bun run build && cd ../..
bun run build:mac
```

The built app will be available in the `release` directory (`.dmg` and `.zip`).

For unsigned local builds:

```bash
bun run build:mac:unsigned
```

### Accessibility Permission

After launching SubMiner for the first time, grant accessibility permission:

1. Open **System Preferences** → **Security & Privacy** → **Privacy** tab
2. Select **Accessibility** from the left sidebar
3. Add SubMiner to the list

Without this permission, window tracking will not work and the overlay won't follow the mpv window.

### macOS Usage Notes

**Launching MPV with IPC:**

```bash
mpv --input-ipc-server=/tmp/subminer-socket video.mkv
```

**Config location:** `$XDG_CONFIG_HOME/SubMiner/config.jsonc` (or `~/.config/SubMiner/config.jsonc`).

**MeCab paths (Homebrew):**

- Apple Silicon (M1/M2): `/opt/homebrew/bin/mecab`
- Intel: `/usr/local/bin/mecab`

Ensure `mecab` is available on your PATH when launching SubMiner.

**Fullscreen:** The overlay should appear correctly in fullscreen. If you encounter issues, check that accessibility permissions are granted.

**mpv plugin binary path:**

```ini
binary_path=/Applications/SubMiner.app/Contents/MacOS/subminer
```

## Windows

Windows support is available through the mpv plugin. Set the binary and socket path in `subminer.conf`:

```ini
binary_path=C:\\Program Files\\subminer\\subminer.exe
socket_path=\\\\.\\pipe\\subminer-socket
```

Launch mpv with:

```bash
mpv --input-ipc-server=\\\\.\\pipe\\subminer-socket video.mkv
```

## MPV Plugin (Optional)

The Lua plugin provides in-player keybindings to control the overlay from mpv. It communicates with SubMiner by invoking the binary with CLI flags.

::: warning Important
mpv must be launched with `--input-ipc-server=/tmp/subminer-socket` for SubMiner to connect.
:::

```bash
# Option 1: install from release assets bundle
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer-assets-0.1.0.tar.gz -O /tmp/subminer-assets.tar.gz
tar -xzf /tmp/subminer-assets.tar.gz -C /tmp
mkdir -p ~/.config/SubMiner
cp /tmp/config.example.jsonc ~/.config/SubMiner/config.jsonc
cp /tmp/plugin/subminer.lua ~/.config/mpv/scripts/
cp /tmp/plugin/subminer.conf ~/.config/mpv/script-opts/

# Option 2: from source checkout
# make install-plugin
```

## Rofi Theme (Optional)

SubMiner ships a default rofi theme at `assets/themes/subminer.rasi`.

Install path (default auto-detected by `subminer`):

- Linux: `~/.local/share/SubMiner/themes/subminer.rasi`
- macOS: `~/Library/Application Support/SubMiner/themes/subminer.rasi`

```bash
mkdir -p ~/.local/share/SubMiner/themes
cp /tmp/assets/themes/subminer.rasi ~/.local/share/SubMiner/themes/subminer.rasi
```

Override with `SUBMINER_ROFI_THEME=/absolute/path/to/theme.rasi`.

All keybindings use a `y` chord prefix — press `y`, then the second key:

| Chord | Action                                |
| ----- | ------------------------------------- |
| `y-y` | Open SubMiner menu (fuzzy-searchable) |
| `y-s` | Start overlay                         |
| `y-S` | Stop overlay                          |
| `y-t` | Toggle visible overlay                |
| `y-i` | Toggle invisible overlay              |
| `y-I` | Show invisible overlay                |
| `y-u` | Hide invisible overlay                |
| `y-o` | Open Yomitan settings                 |
| `y-r` | Restart overlay                       |
| `y-c` | Check overlay status                  |

See [MPV Plugin](/mpv-plugin) for the full configuration reference, script messages, and binary auto-detection details.

## Verify Installation

After installing, confirm SubMiner is working:

```bash
# Start the overlay (connects to mpv IPC)
subminer --start video.mkv

# Useful launch modes for troubleshooting
subminer --log-level debug video.mkv
SubMiner.AppImage --start --log-level debug

# Or with direct AppImage control
SubMiner.AppImage --background  # Background tray service mode
SubMiner.AppImage --start
SubMiner.AppImage --start --dev
SubMiner.AppImage --help    # Show all CLI options
```

You should see the overlay appear over mpv. If subtitles are loaded in the video, they will appear as interactive text in the overlay.

Next: [Usage](/usage) — learn about the `subminer` wrapper, keybindings, and YouTube playback.
