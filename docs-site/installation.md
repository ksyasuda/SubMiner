# Installation

## Requirements

### System Dependencies

| Dependency           | Required   | Notes                                                    |
| -------------------- | ---------- | -------------------------------------------------------- |
| Bun                  | Yes        | Required for `subminer` wrapper and source workflows     |
| mpv                  | Yes        | Must support IPC sockets (`--input-ipc-server`)          |
| ffmpeg               | For media  | Audio extraction and screenshot generation               |
| MeCab + mecab-ipadic | No         | Optional Japanese metadata enrichment (not the primary tokenizer) |
| fuse2                | Linux only | Required for AppImage                                    |
| yt-dlp               | No         | Recommended for YouTube playback and subtitle extraction |

### Platform-Specific

**Linux** — one of the following compositors:

- Hyprland (uses `hyprctl`)
- Sway (uses `swaymsg`)
- X11 (uses `xdotool` and `xwininfo`)

**macOS** — macOS 10.13 or later. Accessibility permission required for window tracking.

**Windows** — Windows 10 or later. Install `mpv` and keep it available on `PATH`; SubMiner's packaged build handles window tracking directly.

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
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/SubMiner.AppImage -O ~/.local/bin/SubMiner.AppImage
chmod +x ~/.local/bin/SubMiner.AppImage

# Download subminer wrapper script
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer -O ~/.local/bin/subminer
chmod +x ~/.local/bin/subminer

```

The `subminer` wrapper uses a Bun shebang (`#!/usr/bin/env bun`), so [Bun](https://bun.sh) must be installed and available on `PATH`.

### From Source

```bash
git clone --recurse-submodules https://github.com/ksyasuda/SubMiner.git
cd SubMiner
# if you cloned without --recurse-submodules:
git submodule update --init --recursive

bun install
bun run build

# Optional packaged Linux artifact
bun run build:appimage
```

Bundled Yomitan is built during `bun run build`.
If you prefer Make wrappers for local install flows, `make build-launcher` still generates `dist/launcher/subminer` and `make install` still installs the wrapper/theme/AppImage when those artifacts exist.

`make build` also builds the bundled Yomitan Chrome extension from the `vendor/subminer-yomitan` submodule into `build/yomitan` using Bun.

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
git clone --recurse-submodules https://github.com/ksyasuda/SubMiner.git
cd SubMiner
git submodule update --init --recursive
make build-macos
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

### Installer (Recommended)

Download the latest Windows installer from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest):

- `SubMiner-<version>.exe` installs the app, Start menu shortcut, and default files under `Program Files`
- `SubMiner-<version>.zip` is available as a portable fallback

Install `mpv` separately and ensure `mpv.exe` is on `PATH`. `ffmpeg` is still required for media extraction, and MeCab remains optional.

### Windows Usage Notes

- Launch `SubMiner.exe` once to let the first-run setup flow seed `%APPDATA%\\SubMiner\\config.jsonc`, offer mpv plugin installation, open bundled Yomitan settings, and optionally create `SubMiner mpv` Start Menu/Desktop shortcuts.
- If you use the mpv plugin, leave `binary_path` empty unless SubMiner is installed in a non-standard location.
- Windows plugin installs rewrite `socket_path` to `\\.\pipe\subminer-socket`; do not keep `/tmp/subminer-socket` on Windows.
- Native window tracking is built in on Windows; no `xdotool`, `xwininfo`, or compositor-specific helper is required.

### From Source (Windows)

```powershell
git clone https://github.com/ksyasuda/SubMiner.git
cd SubMiner
bun install
Set-Location vendor/texthooker-ui
bun install --frozen-lockfile
bun run build
Set-Location ../..
bun run build:win
```

Windows installer builds already get the required NSIS `WinShell` helper through electron-builder's cached `nsis-resources` bundle.
No extra repo-local WinShell plugin install step is required.

## MPV Plugin (Recommended)

The Lua plugin provides in-player keybindings to control the overlay from mpv. It communicates with SubMiner by invoking the binary with CLI flags.

::: warning Important
mpv must be launched with `--input-ipc-server=/tmp/subminer-socket` for SubMiner to connect.
:::

On Windows, the packaged plugin config is rewritten to `socket_path=\\.\pipe\subminer-socket`.

```bash
# Option 1: install from release assets bundle
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer-assets.tar.gz -O /tmp/subminer-assets.tar.gz
tar -xzf /tmp/subminer-assets.tar.gz -C /tmp
mkdir -p ~/.config/SubMiner
cp /tmp/config.example.jsonc ~/.config/SubMiner/config.jsonc
mkdir -p ~/.config/mpv/scripts/subminer
mkdir -p ~/.config/mpv/script-opts
cp -R /tmp/plugin/subminer/. ~/.config/mpv/scripts/subminer/
cp /tmp/plugin/subminer.conf ~/.config/mpv/script-opts/

# Option 2: from source checkout
# make install-plugin
```

## Rofi Theme (Linux Only)

SubMiner ships a default rofi theme at `assets/themes/subminer.rasi`.

Install path (default auto-detected by `subminer`):

- Linux: `~/.local/share/SubMiner/themes/subminer.rasi`
- macOS: `~/Library/Application Support/SubMiner/themes/subminer.rasi`

```bash
mkdir -p ~/.local/share/SubMiner/themes
cp /tmp/assets/themes/subminer.rasi ~/.local/share/SubMiner/themes/subminer.rasi
```

Override with `SUBMINER_ROFI_THEME=/absolute/path/to/theme.rasi`.

See [MPV Plugin](/mpv-plugin) for the full configuration reference, keybindings, script messages, and binary auto-detection details.

## Verify Installation

After installing, confirm SubMiner is working:

On Windows, replace `SubMiner.AppImage` with `SubMiner.exe` in the direct app commands below.

```bash
# Play a file (default plugin config auto-starts visible overlay and waits for annotation readiness; first launch may open first-run setup popup)
subminer video.mkv

# Optional explicit overlay start for setups with plugin auto_start=no
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
