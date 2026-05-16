# Installation

## How the Pieces Fit Together

SubMiner is an overlay that sits on top of mpv. It connects to mpv through an IPC socket, renders subtitles as interactive text using a bundled Yomitan dictionary engine, and optionally creates Anki flashcards via AnkiConnect.

To get a working setup you need:

1. **mpv** launched with an IPC socket so SubMiner can read subtitle data
2. **SubMiner** (the Electron overlay app)
3. **Dictionaries** imported into the bundled Yomitan instance (lookups won't work without at least one)
4. **Anki + AnkiConnect** _(optional but recommended)_ for card creation and enrichment

The `subminer` launcher script handles step 1 automatically. If you launch mpv yourself or from another tool, you must pass `--input-ipc-server=/tmp/subminer-socket` (or the equivalent named pipe on Windows) — without it the overlay will start but subtitles will never appear.

## Requirements

### System Dependencies

| Dependency           | Required    | Notes                                                                                                                                                                          |
| -------------------- | ----------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| mpv                  | Yes         | Must support IPC sockets (`--input-ipc-server`)                                                                                                                                |
| Bun                  | For wrapper | Required for `subminer` CLI wrapper and source builds. Pre-built releases (AppImage, DMG, installer) work without it — only the `subminer` wrapper script needs Bun on `PATH`. |
| ffmpeg               | Recommended | Audio extraction and screenshot generation. Without it SubMiner still runs, but audio and image fields on Anki cards will be empty.                                            |
| MeCab + mecab-ipadic | No          | Adds part-of-speech data used to filter particles out of N+1, JLPT, and frequency annotations. Without it annotations still render, but POS-based filtering is less precise.   |
| fuse2                | Linux only  | Required for AppImage                                                                                                                                                          |
| yt-dlp               | No          | Recommended for YouTube playback and subtitle extraction                                                                                                                       |

### Platform-Specific

**Linux** — one of the following window backends:

- **Hyprland** — native Wayland support (uses `hyprctl`)
- **Sway** — native Wayland support (uses `swaymsg`)
- **X11 / Xwayland** — for X11 sessions or any other Wayland compositor (uses `xdotool` and `xwininfo`)

::: warning Wayland support is compositor-specific
Wayland has no universal API for window positioning — each compositor exposes its own IPC, so SubMiner needs a dedicated backend per compositor. Only Hyprland and Sway have native Wayland backends. If you run a different Wayland compositor (GNOME, KDE Plasma, river, etc.), both mpv **and** SubMiner must run under X11 or Xwayland. The `subminer` launcher handles this automatically when `--backend x11` is set or the X11 backend is auto-detected.
:::

<details>
<summary><b>Arch Linux</b></summary>

```bash
sudo pacman -S --needed mpv ffmpeg
# Optional
sudo pacman -S --needed mecab mecab-ipadic yt-dlp fzf rofi chafa ffmpegthumbnailer
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
# Optional
sudo apt install mecab libmecab-dev mecab-ipadic-utf8 fzf rofi chafa ffmpegthumbnailer yt-dlp
# X11 / Xwayland (required for non-Hyprland/Sway compositors)
sudo apt install xdotool x11-utils
# Optional: subtitle sync
pip install ffsubsync
# alass is not in apt — install via cargo: cargo install alass-cli
```

</details>

<details>
<summary><b>Fedora</b></summary>

```bash
sudo dnf install mpv ffmpeg
# Optional
sudo dnf install mecab mecab-ipadic fzf rofi chafa ffmpegthumbnailer yt-dlp
# X11 / Xwayland (required for non-Hyprland/Sway compositors)
sudo dnf install xdotool xorg-x11-utils
# Optional: subtitle sync
pip install ffsubsync
# alass: cargo install alass-cli
```

</details>

**macOS** — macOS 10.13 or later. Accessibility permission required for window tracking.

```bash
brew install mpv ffmpeg
# Optional but recommended for annotations
brew install mecab mecab-ipadic
# Optional
brew install yt-dlp fzf rofi chafa ffmpegthumbnailer
# Optional: subtitle sync
brew install alass
pip install ffsubsync
```

**Windows** — Windows 10 or later. Install [`mpv`](https://mpv.io/installation/) and [`ffmpeg`](https://ffmpeg.org/download.html) and ensure both are on `PATH`. Keep `mpv.exe` on `PATH` for auto-discovery or set `mpv.executablePath` in config if it lives elsewhere. SubMiner's packaged build handles window tracking directly. Optionally install [MeCab for Windows](https://taku910.github.io/mecab/#download) with the UTF-8 dictionary.

### Optional Tools

| Tool              | Purpose                                                                                                                                          |
| ----------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| fzf               | Terminal-based video picker (default)                                                                                                            |
| rofi              | GUI-based video picker                                                                                                                           |
| chafa             | Thumbnail previews in fzf                                                                                                                        |
| ffmpegthumbnailer | Generate video thumbnails for picker                                                                                                             |
| guessit           | Better AniSkip title/season/episode parsing for file playback                                                                                    |
| alass             | Subtitle sync engine (preferred) — must be on `PATH` or set `subsync.alass_path` in config; subtitle syncing is disabled without it or ffsubsync |
| ffsubsync         | Subtitle sync engine (fallback) — must be on `PATH` or set `subsync.ffsubsync_path` in config; subtitle syncing is disabled without it or alass  |

## Linux

### Arch Linux (AUR)

Install [`subminer-bin`](https://aur.archlinux.org/packages/subminer-bin) from the AUR if you want the packaged Linux release managed by pacman. The package installs the official SubMiner AppImage plus the `subminer` wrapper.

```bash
paru -S subminer-bin
```

Or manually:

```bash
git clone https://aur.archlinux.org/subminer-bin.git
cd subminer-bin
makepkg -si
```

### AppImage (Recommended)

Download the latest AppImage and the `subminer` launcher from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest).

**Step 1 — Install Bun** (required for the launcher):

```bash
curl -fsSL https://bun.sh/install | bash
```

The `subminer` launcher uses a Bun shebang. The AppImage itself does **not** need Bun — only the launcher does. If you skip the launcher and run the AppImage directly (for example `SubMiner.AppImage --start`), you can skip this step, but you will need to configure `mpv.conf` with `input-ipc-server=/tmp/subminer-socket` manually.

**Step 2 — Download and install:**

```bash
mkdir -p ~/.local/bin

# Download and install AppImage
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/SubMiner.AppImage -O ~/.local/bin/SubMiner.AppImage
chmod +x ~/.local/bin/SubMiner.AppImage

# Download and install the subminer launcher (recommended)
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer -O ~/.local/bin/subminer
chmod +x ~/.local/bin/subminer

# Download launcher support assets used for bundled runtime plugin injection
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer-assets.tar.gz -O /tmp/subminer-assets.tar.gz
tar -xzf /tmp/subminer-assets.tar.gz -C /tmp
mkdir -p ~/.local/share/SubMiner/plugin/subminer
cp -R /tmp/plugin/subminer/. ~/.local/share/SubMiner/plugin/subminer/
```

The `subminer` launcher is the recommended way to use SubMiner on Linux. It ensures mpv is launched with the correct IPC socket, SubMiner defaults, and the bundled runtime plugin so you don't need to configure `mpv.conf` or install a global mpv plugin.

The first-run setup window can also install Bun and the packaged `subminer` launcher into an existing writable PATH directory. Both steps are optional.

To check for updates later:

```bash
subminer -u
# or
subminer --update
```

SubMiner verifies launcher/support asset downloads against `SHA256SUMS.txt`. If the launcher is installed in a protected path such as `/usr/local/bin/subminer`, SubMiner does not elevate itself; it shows the exact `sudo curl ... && sudo chmod +x ...` command to run instead.

On Linux, `subminer -u` performs this update from the launcher process, so it does not need to start or IPC into the tray app.

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
brew install mpv ffmpeg
# Optional but recommended if you use N+1, JLPT, or frequency annotations
brew install mecab mecab-ipadic
```

#### Install the `subminer` launcher (recommended)

The `subminer` launcher is the recommended way to use SubMiner on macOS. It launches mpv with the correct IPC socket and SubMiner defaults so you don't need to set up an `mpv.conf` profile manually.

First-run setup can install Bun and the packaged launcher into a writable directory that is already on PATH. It does not edit shell profiles.

Download it from the same [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest) page:

```bash
sudo wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer -O /usr/local/bin/subminer
sudo chmod +x /usr/local/bin/subminer
```

Or with curl:

```bash
sudo curl -fSL https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer -o /usr/local/bin/subminer
sudo chmod +x /usr/local/bin/subminer
```

To check for updates later:

```bash
subminer -u
# or
subminer --update
```

SubMiner verifies launcher/support asset downloads against `SHA256SUMS.txt`. If `/usr/local/bin/subminer` is protected, SubMiner shows the exact `sudo curl ... && sudo chmod +x ...` command to run instead of elevating itself.

::: warning Bun required for the launcher
The `subminer` launcher uses a Bun shebang (`#!/usr/bin/env bun`), so [Bun](https://bun.sh) must be installed and available on `PATH`. Install Bun if you haven't already: `curl -fsSL https://bun.sh/install | bash`.
:::

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

Build and install the launcher alongside the app:

```bash
make install-macos
```

This builds the `subminer` launcher into `dist/launcher/subminer` and installs it to `~/.local/bin/subminer` along with the app bundle and rofi theme. To install to `/usr/local/bin` instead (already on the default macOS `PATH`):

```bash
sudo make install-macos PREFIX=/usr/local
```

### Gatekeeper

If macOS blocks SubMiner on first launch, right-click the app and select **Open** to bypass the warning. Alternatively, remove the quarantine attribute:

```bash
xattr -d com.apple.quarantine /Applications/SubMiner.app
```

### Accessibility Permission

After launching SubMiner for the first time, grant accessibility permission:

1. Open **System Settings** → **Privacy & Security** → **Accessibility**
2. Enable SubMiner in the list (add it if it does not appear)

Without this permission, window tracking will not work and the overlay won't follow the mpv window.

### macOS Usage Notes

**Launching with the `subminer` launcher (recommended):**

```bash
subminer video.mkv
```

The launcher handles the IPC socket and SubMiner defaults automatically. If you prefer to launch mpv manually:

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

### Prerequisites

1. Install [`mpv`](https://mpv.io/installation/) and ensure `mpv.exe` is on `PATH`. If mpv is installed elsewhere, you can set `mpv.executablePath` in `config.jsonc` or use the first-run setup field to point at the executable.
2. Install [`ffmpeg`](https://ffmpeg.org/download.html) and add it to `PATH` — recommended for audio/screenshot extraction (without it, media fields on Anki cards will be empty).
3. _(Optional)_ Install [MeCab for Windows](https://taku910.github.io/mecab/#download) with the UTF-8 dictionary for annotation POS filtering.

No compositor tools or window helpers are needed — native window tracking is built in on Windows.

### Installer (Recommended)

Download the latest Windows installer from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest):

- `SubMiner-<version>.exe` installs the app, Start menu shortcut, and default files under `Program Files`
- `SubMiner-<version>.zip` is available as a portable fallback

### Getting Started on Windows

1. **Run `SubMiner.exe` once** — first-run setup creates `%APPDATA%\SubMiner\config.jsonc` and opens Yomitan settings for dictionary import. The global mpv plugin install is optional for compatibility; the SubMiner mpv shortcut injects the bundled runtime plugin.
2. **Create the SubMiner mpv shortcut** _(recommended)_ — the setup popup offers to create a `SubMiner mpv` Start Menu and/or Desktop shortcut. This is the recommended way to launch playback on Windows.
3. **Optional: install the command-line launcher** — first-run setup can install Bun with winget/Scoop/the official installer and add `%LOCALAPPDATA%\SubMiner\bin\subminer.cmd` to your user PATH. Open a new terminal and type `subminer`.
4. **Play a video** — double-click the shortcut, drag a video file onto it, or run from a terminal:

```powershell
& "C:\Program Files\SubMiner\SubMiner.exe" --launch-mpv "C:\Videos\episode 01.mkv"
```

The shortcut and `--launch-mpv` pass SubMiner's default IPC socket, subtitle args, and bundled runtime plugin directly — no `mpv.conf` profile or global mpv plugin install is needed.

### Windows-Specific Notes

- The **SubMiner mpv** shortcut created during first-run setup is the recommended way to launch playback on Windows.
- The optional command-line launcher installs a `subminer.cmd` shim, but users type `subminer`; Windows resolves `.cmd` through `PATHEXT`.
- First-run setup adds only `%LOCALAPPDATA%\SubMiner\bin` to the HKCU user PATH. It does not add `SubMiner.exe` or the app install directory to PATH.
- First-run plugin installs pin `binary_path` to the current `SubMiner.exe` automatically. Manual plugin configs can leave `binary_path` empty unless SubMiner is in a non-standard location.
- Plugin installs rewrite `socket_path` to `\\.\pipe\subminer-socket` — do not keep `/tmp/subminer-socket` on Windows.
- Config is stored at `%APPDATA%\SubMiner\config.jsonc`.

### From Source (Windows)

```powershell
git clone https://github.com/ksyasuda/SubMiner.git
cd SubMiner
bun install

# Windows requires building the texthooker-ui submodule manually before
# the main build (Linux/macOS handle this automatically during `bun run build`).
Set-Location vendor/texthooker-ui
bun install --frozen-lockfile
bun run build
Set-Location ../..

bun run build:win
```

Windows installer builds already get the required NSIS `WinShell` helper through electron-builder's cached `nsis-resources` bundle.
No extra repo-local WinShell plugin install step is required.

## MPV Plugin

SubMiner-managed playback loads the bundled mpv plugin at runtime. No separate global mpv plugin install is required when launching from the app, the launcher, or the packaged Windows SubMiner mpv shortcut.

::: warning Important
If first-run setup detects an older global SubMiner mpv plugin under mpv's `scripts` directory, use **Remove legacy mpv plugin** so regular mpv playback stops loading SubMiner.
:::

See [MPV Plugin](/mpv-plugin) for the keybindings, script messages, and runtime configuration reference.

## Anki Setup (Recommended)

If you plan to mine Anki cards (the primary use case for most users):

1. Install [Anki](https://apps.ankiweb.net/).
2. Install the [AnkiConnect](https://ankiweb.net/shared/info/2055492159) add-on — open Anki, go to **Tools → Add-ons → Get Add-ons**, enter code `2055492159`.
3. Restart Anki and keep it running while using SubMiner.

AnkiConnect listens on `http://127.0.0.1:8765` by default. SubMiner will connect to it automatically with no extra config needed for basic card creation.

For enrichment configuration (sentence, audio, screenshot fields), see [Anki Integration](/anki-integration).

## First-Run Setup

Run the setup wizard to create a default config and finish initial configuration. You do **not** need to create the config manually — SubMiner handles it.

```bash
subminer app --setup
```

> [!NOTE]
> On Windows, run `SubMiner.exe` directly — it opens the setup wizard automatically on first launch.

The setup popup walks you through:

- **Config file**: auto-created at `~/.config/SubMiner/config.jsonc` (Linux/macOS) or `%APPDATA%\SubMiner\config.jsonc` (Windows)
- **mpv plugin**: install the bundled Lua plugin for in-player keybindings
- **Yomitan dictionaries**: import at least one dictionary so lookups work
- **Windows shortcut** _(Windows only)_: optionally create a `SubMiner mpv` Start Menu/Desktop shortcut
- **Command line launcher**: optionally install Bun and the `subminer` launcher to your command-line PATH

The `Finish setup` button follows the normal config/Yomitan readiness checks. Bun and the command-line launcher are optional and never block setup completion.

> [!TIP]
> You can re-open the setup popup at any time with `subminer app --setup` or `SubMiner.AppImage --setup`.

Once setup is complete, play a video to verify everything works:

```bash
subminer video.mkv
```

You should see the overlay appear over mpv. If subtitles are loaded in the video, they will appear as interactive text in the overlay.

<details>
<summary><b>More launch examples</b></summary>

```bash
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

</details>

## Verify Setup

After completing first-run setup, run the built-in diagnostic to confirm everything is in place:

```bash
subminer doctor
```

This checks for the app binary, mpv, ffmpeg, config file, and socket path. Fix any failures before continuing.

> [!NOTE]
> On Windows, run `SubMiner.exe` directly. Replace `SubMiner.AppImage` with `SubMiner.exe` in the direct app commands below.

## Optional Extras

### Rofi Theme (Linux Only)

SubMiner ships a custom rofi theme bundled in the release assets tarball.

Install path (default auto-detected by `subminer`):

- Linux: `~/.local/share/SubMiner/themes/subminer.rasi`
- macOS: `~/Library/Application Support/SubMiner/themes/subminer.rasi`

```bash
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer-assets.tar.gz -O /tmp/subminer-assets.tar.gz
tar -xzf /tmp/subminer-assets.tar.gz -C /tmp
mkdir -p ~/.local/share/SubMiner/themes
cp /tmp/assets/themes/subminer.rasi ~/.local/share/SubMiner/themes/subminer.rasi
```

Override with `SUBMINER_ROFI_THEME=/absolute/path/to/theme.rasi`.

Next: [Usage](/usage) — learn about the `subminer` wrapper, keybindings, and YouTube playback.
