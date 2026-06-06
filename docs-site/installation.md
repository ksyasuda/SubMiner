# Installation

SubMiner is a desktop app that draws an interactive layer - an **overlay** - on top of the [mpv](https://mpv.io) video player. As you watch native Japanese media, you can click or hover any word in the subtitles to look it up, then turn it into an Anki flashcard without pausing to switch apps. Building flashcards from real content you're watching is called **sentence mining**, and it's what SubMiner is built for. It bundles its own copy of **Yomitan** (a pop-up dictionary) and talks to **AnkiConnect** (an add-on that lets other programs add cards to Anki) so cards get filled in automatically.

Three steps to get started:

1. **Install requirements** - mpv and a few optional extras
2. **Install SubMiner** - from the AUR, or download from GitHub Releases
3. **Launch the app** - first-run setup walks you through dictionaries, the launcher, and everything else

## 1. Install Requirements

Only **mpv** is strictly required to run SubMiner. Everything else enhances the experience but is optional.

| Dependency           | Status      | What it does                                                                                                    |
| -------------------- | ----------- | --------------------------------------------------------------------------------------------------------------- |
| mpv                  | Required    | The video player SubMiner overlays on. Must support `--input-ipc-server`.                                       |
| ffmpeg               | Recommended | Audio extraction and screenshots for Anki cards. Without it SubMiner still runs, but media fields will be empty. |
| MeCab + mecab-ipadic | Recommended | Part-of-speech filtering for more precise N+1, JLPT, and frequency annotations. Without it annotations still render, but POS-based filtering is less accurate. |
| yt-dlp               | Optional    | YouTube playback and subtitle extraction.                                                                       |
| fzf                  | Optional    | Terminal-based video picker in the launcher.                                                                     |
| rofi                 | Optional    | GUI-based video picker (Linux).                                                                                  |
| chafa                | Optional    | Thumbnail previews in fzf.                                                                                       |
| ffmpegthumbnailer    | Optional    | Video thumbnail generation for the picker.                                                                       |
| guessit              | Optional    | Better AniSkip title/season/episode parsing.                                                                     |
| alass                | Optional    | Subtitle sync engine (preferred). Disabled without alass or ffsubsync.                                           |
| ffsubsync            | Optional    | Audio-based subtitle sync engine. Disabled without alass or ffsubsync.                                           |
| fuse2                | Linux only  | Required to run the AppImage.                                                                                    |

### Linux

**Window backend** - you need one of these depending on your compositor:

- **Hyprland** - native Wayland support (uses `hyprctl`)
- **Sway** - native Wayland support (uses `swaymsg`)
- **X11 / Xwayland** - for X11 sessions or any other Wayland compositor (uses `xdotool` and `xwininfo`)

::: warning Wayland support is compositor-specific
Wayland has no universal API for window positioning - each compositor exposes its own IPC, so SubMiner needs a dedicated backend per compositor. Only Hyprland and Sway have native Wayland backends. If you run a different Wayland compositor (GNOME, KDE Plasma, river, etc.), both mpv **and** SubMiner must run under X11 or Xwayland. The `subminer` launcher handles this automatically when `--backend x11` is set or the X11 backend is auto-detected.
:::

<details>
<summary><b>Arch Linux</b></summary>

```bash
sudo pacman -S --needed mpv ffmpeg
# Recommended
sudo pacman -S --needed mecab mecab-ipadic
# Optional
sudo pacman -S --needed yt-dlp fzf rofi chafa ffmpegthumbnailer
# Optional: subtitle sync (at least one needed for subtitle syncing)
paru -S --needed alass python-ffsubsync
# X11 / Xwayland (required for non-Hyprland/Sway compositors)
sudo pacman -S --needed xdotool xorg-xwininfo
```

</details>

<details>
<summary><b>Ubuntu / Debian</b></summary>

```bash
sudo apt install mpv ffmpeg
# Recommended
sudo apt install mecab libmecab-dev mecab-ipadic-utf8
# Optional
sudo apt install yt-dlp fzf rofi chafa ffmpegthumbnailer
# X11 / Xwayland (required for non-Hyprland/Sway compositors)
sudo apt install xdotool x11-utils
# Optional: subtitle sync
pip install ffsubsync
# alass is not in apt - install via cargo: cargo install alass-cli
```

</details>

<details>
<summary><b>Fedora</b></summary>

```bash
sudo dnf install mpv ffmpeg
# Recommended
sudo dnf install mecab mecab-ipadic
# Optional
sudo dnf install yt-dlp fzf rofi chafa ffmpegthumbnailer
# X11 / Xwayland (required for non-Hyprland/Sway compositors)
sudo dnf install xdotool xorg-x11-utils
# Optional: subtitle sync
pip install ffsubsync
# alass: cargo install alass-cli
```

</details>

### macOS

macOS 11 (Big Sur) or later. Accessibility permission - the macOS setting that lets one app observe and position another app's windows - is required so the overlay can follow the mpv window (see [step 2](#macos-dmg)).

```bash
brew install mpv ffmpeg
# Recommended
brew install mecab mecab-ipadic
# Optional
brew install yt-dlp fzf chafa ffmpegthumbnailer
# Optional: subtitle sync
brew install alass
pip install ffsubsync
```

### Windows

Windows 10 or later. Install [`mpv`](https://mpv.io/installation/) and [`ffmpeg`](https://ffmpeg.org/download.html) and ensure both are on `PATH`. Optionally install [MeCab for Windows](https://taku910.github.io/mecab/#download) with the UTF-8 dictionary.

No compositor tools or window helpers are needed - native window tracking is built in.

## 2. Install SubMiner

### Arch Linux (AUR) {#arch-aur}

Install [`subminer-bin`](https://aur.archlinux.org/packages/subminer-bin) from the AUR. The package includes the SubMiner AppImage and the `subminer` launcher.

```bash
paru -S subminer-bin
```

Or manually:

```bash
git clone https://aur.archlinux.org/subminer-bin.git
cd subminer-bin
makepkg -si
```

### Linux (AppImage) {#linux-appimage}

Download the latest AppImage from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest):

```bash
mkdir -p ~/.local/bin
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/SubMiner.AppImage -O ~/.local/bin/SubMiner.AppImage
chmod +x ~/.local/bin/SubMiner.AppImage
```

::: tip Launcher install is optional
First-run setup can install [Bun](https://bun.sh) and the `subminer` command-line launcher for you automatically. You don't need to download the launcher separately.

If you prefer to install it manually, see [manual launcher install](#manual-launcher-install-linux).
:::

### macOS (DMG) {#macos-dmg}

Download the DMG from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest), open it, and drag `SubMiner.app` into `/Applications`. A ZIP artifact is also available as a fallback.

**Gatekeeper:** If macOS blocks SubMiner on first launch, right-click the app and select **Open** to bypass the warning. Alternatively:

```bash
xattr -d com.apple.quarantine /Applications/SubMiner.app
```

**Accessibility permission:** Grant accessibility permission so the overlay can track the mpv window:

1. Open **System Settings** → **Privacy & Security** → **Accessibility**
2. Enable SubMiner in the list (add it if it does not appear)

::: tip Launcher install is optional
First-run setup can install [Bun](https://bun.sh) and the `subminer` command-line launcher for you automatically. You don't need to download the launcher separately.

If you prefer to install it manually, see [manual launcher install](#manual-launcher-install-macos).
:::

### Windows (Installer) {#windows-installer}

Download the latest installer from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest):

- `SubMiner-<version>.exe` - installer (recommended)
- `SubMiner-<version>-win.zip` - portable fallback

Make sure `mpv.exe` is on your `PATH`, or set `mpv.executablePath` in the config during first-run setup.

### From Source

<details>
<summary><b>Linux</b></summary>

```bash
git clone --recurse-submodules https://github.com/ksyasuda/SubMiner.git
cd SubMiner
make deps
bun run build

# Optional: build AppImage
bun run build:appimage
```

Bundled Yomitan is built during `bun run build`.

</details>

<details>
<summary><b>macOS</b></summary>

```bash
git clone --recurse-submodules https://github.com/ksyasuda/SubMiner.git
cd SubMiner
make deps
make build-macos
```

The built app will be in the `release` directory (`.dmg` and `.zip`). For unsigned local builds: `bun run build:mac:unsigned`.

</details>

<details>
<summary><b>Windows</b></summary>

```powershell
git clone https://github.com/ksyasuda/SubMiner.git
cd SubMiner
git submodule update --init --recursive
bun install
Set-Location stats
bun install --frozen-lockfile
Set-Location ../vendor/texthooker-ui
bun install --frozen-lockfile
bun run build
Set-Location ../..
bun run build:win
```

</details>

## 3. Launch & First-Run Setup

Launch SubMiner and the setup wizard will open automatically:

```bash
# Linux (AUR install)
subminer app --setup

# Linux (AppImage directly)
~/.local/bin/SubMiner.AppImage --setup

# macOS - launch SubMiner.app from /Applications, or:
subminer app --setup
```

On **Windows**, just run `SubMiner.exe` - the setup wizard opens automatically on first launch.

The setup wizard walks you through:

- **Config file** - auto-created at `~/.config/SubMiner/config.jsonc` (Linux/macOS) or `%APPDATA%\SubMiner\config.jsonc` (Windows)
- **Yomitan dictionaries** - import at least one dictionary so word lookups work
- **Bun + `subminer` launcher** _(optional)_ - installs the command-line launcher into a writable PATH directory
- **Windows shortcut** _(Windows only)_ - create a `SubMiner mpv` Start Menu/Desktop shortcut

The `Finish setup` button requires a config file and at least one Yomitan dictionary. Bun and the launcher are optional and never block setup completion.

> [!TIP]
> You can re-open the setup wizard at any time with `subminer app --setup` or `SubMiner.AppImage --setup`.

### Play a Video

Once setup is complete:

```bash
subminer video.mkv
```

You should see the overlay appear over mpv. If subtitles are loaded, they will appear as interactive text in the overlay.

On **Windows**, the recommended way to play video is with the **SubMiner mpv** shortcut created during setup - double-click it, or drag a video file onto it.

### Verify Setup

Run the built-in diagnostic to confirm everything is working:

```bash
subminer doctor
```

This checks for the app binary, mpv, ffmpeg, config file, and socket path. Fix any failures before continuing.

## Anki Setup (Recommended)

If you plan to mine Anki cards:

1. Install [Anki](https://apps.ankiweb.net/)
2. Install [AnkiConnect](https://ankiweb.net/shared/info/2055492159) - open Anki → **Tools → Add-ons → Get Add-ons** → enter code `2055492159`
3. Restart Anki and keep it running while using SubMiner

AnkiConnect listens on `http://127.0.0.1:8765` by default. SubMiner connects automatically with no extra config needed.

For enrichment configuration (sentence, audio, screenshot fields), see [Anki Integration](/anki-integration).

## Updates

```bash
subminer -u
# or
subminer --update
```

SubMiner verifies AppImage, launcher, and rofi theme downloads against `SHA256SUMS.txt`. If the binary is in a protected path, SubMiner shows the exact command to run rather than elevating itself.

The tray "Check for Updates" entry installs the new app automatically on Linux, macOS, and Windows. On Linux it replaces the running `.AppImage` in place via `electron-updater`; AppImages managed by a system package (for example the AUR `/opt/SubMiner/SubMiner.AppImage`) are skipped so the package manager stays in charge.

`subminer -u` also performs the AppImage update directly from the launcher process, which is useful when SubMiner is not currently running.

## How It All Fits Together

SubMiner is an overlay that sits on top of mpv. It connects to mpv through an IPC socket, renders subtitles as interactive text using a bundled Yomitan dictionary engine, and optionally creates Anki flashcards via AnkiConnect.

The `subminer` launcher handles mpv IPC socket setup automatically. If you launch mpv yourself or from another tool, you must pass `--input-ipc-server=/tmp/subminer-socket` (or `\\.\pipe\subminer-socket` on Windows) - without it the overlay starts but subtitles won't appear.

The bundled mpv plugin is injected at runtime automatically - you don't need to install it separately. It provides in-player keybindings (the `y` chord) for controlling the overlay from within mpv. See [MPV Plugin](/mpv-plugin) for the full keybinding and configuration reference.

## Platform Notes

### macOS

**MeCab paths (Homebrew):**
- Apple Silicon (M1/M2): `/opt/homebrew/bin/mecab`
- Intel: `/usr/local/bin/mecab`

Ensure `mecab` is available on your PATH when launching SubMiner.

**Fullscreen:** The overlay should appear correctly in fullscreen. If you encounter issues, check that accessibility permissions are granted.

### Windows

- The **SubMiner mpv** shortcut is the recommended way to launch playback. It starts `mpv.exe` with the right IPC socket and subtitle defaults.
- First-run setup adds only `%LOCALAPPDATA%\SubMiner\bin` to the HKCU user PATH. It does not add `SubMiner.exe` to PATH.
- IPC socket on Windows is `\\.\pipe\subminer-socket` - do not use `/tmp/subminer-socket`.
- Config is stored at `%APPDATA%\SubMiner\config.jsonc`.

## Manual Launcher Install

The `subminer` launcher uses a [Bun](https://bun.sh) shebang, so Bun must be installed. First-run setup can handle this automatically, but if you prefer to do it yourself:

### Linux {#manual-launcher-install-linux}

```bash
# Install Bun
curl -fsSL https://bun.sh/install | bash

# Download the launcher
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer -O ~/.local/bin/subminer
chmod +x ~/.local/bin/subminer
```

### macOS {#manual-launcher-install-macos}

```bash
# Install Bun
curl -fsSL https://bun.sh/install | bash

# Download the launcher
sudo curl -fSL https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer -o /usr/local/bin/subminer
sudo chmod +x /usr/local/bin/subminer
```

## Optional Extras

### Rofi Theme (Linux Only)

SubMiner ships a custom rofi theme in the release assets:

```bash
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer-assets.tar.gz -O /tmp/subminer-assets.tar.gz
tar -xzf /tmp/subminer-assets.tar.gz -C /tmp
mkdir -p ~/.local/share/SubMiner/themes
cp /tmp/assets/themes/subminer.rasi ~/.local/share/SubMiner/themes/subminer.rasi
```

Override with `SUBMINER_ROFI_THEME=/absolute/path/to/theme.rasi`.

Next: [Usage](/usage) - learn about the `subminer` wrapper, keybindings, and YouTube playback.
