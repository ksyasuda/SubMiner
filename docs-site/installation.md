# Installation

SubMiner is a desktop app that draws an interactive layer - an **overlay** - on top of the [mpv](https://mpv.io) video player. As you watch native Japanese media, you can click or hover any word in the subtitles to look it up, then turn it into an Anki flashcard without pausing to switch apps. Building flashcards from real content you're watching is called **sentence mining**, and it's what SubMiner is built for. It bundles its own copy of **Yomitan** (a pop-up dictionary) and talks to **AnkiConnect** (an add-on that lets other programs add cards to Anki) so cards get filled in automatically.

Three steps to get started:

1. **Install requirements** - mpv and a few optional extras
2. **Install SubMiner** - from the AUR, or download from GitHub Releases
3. **Launch the app** - first-run setup walks you through dictionaries, the launcher, and everything else

## 1. Install Requirements

Only **mpv** is strictly required to run SubMiner. Everything else enhances the experience but is optional.

Several entries below exist only for the `subminer` command-line launcher, which is Linux and macOS only. On Windows you launch playback with the **SubMiner mpv** shortcut instead, so you can ignore those rows.

| Dependency           | Status      | Platforms    | What it does                                                                                                                                                   |
| -------------------- | ----------- | ------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| mpv                  | Required    | All          | The video player SubMiner overlays on. Must support `--input-ipc-server`.                                                                                      |
| ffmpeg               | Recommended | All          | Audio extraction and screenshots for Anki cards. Without it SubMiner still runs, but media fields will be empty.                                               |
| MeCab + mecab-ipadic | Recommended | All          | Part-of-speech filtering for more precise N+1, JLPT, and frequency annotations. Without it annotations still render, but POS-based filtering is less accurate. |
| yt-dlp               | Optional    | All          | YouTube playback and subtitle extraction.                                                                                                                      |
| xz                   | Optional    | All          | Required for TsukiHime subtitle downloads (subtitles are served xz-compressed). Preinstalled on most Linux distros; not present on Windows by default.         |
| guessit              | Optional    | All          | Better AniSkip title/season/episode parsing.                                                                                                                   |
| alass                | Optional    | All          | Subtitle sync engine (preferred). Disabled without alass or ffsubsync.                                                                                         |
| ffsubsync            | Optional    | All          | Audio-based subtitle sync engine. Disabled without alass or ffsubsync.                                                                                         |
| fzf                  | Optional    | Linux, macOS | Terminal-based video picker in the `subminer` launcher.                                                                                                        |
| rofi                 | Optional    | Linux        | GUI-based video picker in the `subminer` launcher.                                                                                                             |
| chafa                | Optional    | Linux, macOS | Thumbnail previews in the fzf picker.                                                                                                                          |
| ffmpegthumbnailer    | Optional    | Linux, macOS | Video thumbnail generation for the pickers.                                                                                                                    |
| fuse2                | Required    | Linux        | Needed to run the AppImage.                                                                                                                                    |

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

Windows 10 or later. No compositor tools or window helpers are needed - native window tracking is built in.

You need **mpv** (required) and **ffmpeg** (strongly recommended, for card audio and screenshots), and both must be on your `PATH`.

::: tip What is PATH?
`PATH` is the list of folders Windows searches when a program asks to run another program by name. SubMiner runs `mpv` and `ffmpeg` by name, so if their folders are not on `PATH`, SubMiner cannot find them even though they are installed. The routes below mostly handle `PATH` for you; the manual route explains how to add a folder yourself.
:::

You can install these with a package manager or by hand. Coverage differs, so pick based on what you need:

| Dependency       | winget          | Scoop         |
| ---------------- | --------------- | ------------- |
| mpv (required)   | `shinchiro.mpv` | `extras/mpv`  |
| ffmpeg           | `Gyan.FFmpeg`   | `main/ffmpeg` |
| yt-dlp (YouTube) | `yt-dlp.yt-dlp` | `main/yt-dlp` |
| xz (TsukiHime)   | not packaged    | `main/xz`     |

Use **winget** if you want Microsoft's first-party tool and don't need TsukiHime subtitle downloads. Use **Scoop** if you want one package manager to cover everything, since it is the only one that also packages `xz`.

#### Recommended: winget

[winget](https://learn.microsoft.com/windows/package-manager/winget/) is Microsoft's own package manager and ships with Windows 11 and current Windows 10 (it comes with **App Installer** from the Microsoft Store). In **PowerShell** or **Command Prompt**:

```powershell
winget install shinchiro.mpv
winget install Gyan.FFmpeg
```

Close and reopen your terminal, then check that both are found:

```powershell
mpv --version
ffmpeg -version
```

`ffmpeg` is installed as a portable package, so winget links it into a folder that is already on your `PATH` and it should work right away.

`mpv` uses a regular installer, and depending on the version it may **not** add itself to `PATH`. If `mpv --version` says `not recognized`, you have two easy options:

- Note where it installed (usually `%LOCALAPPDATA%\Programs\mpv`) and add that folder to `PATH` using the manual steps below, or
- Skip `PATH` entirely and set `mpv.executablePath` to the full path of `mpv.exe` during first-run setup.

Once `mpv --version` works, or you have the full path to `mpv.exe` ready, continue to [step 2](#_2-install-subminer).

<details>
<summary><b>Alternative: Scoop (covers every dependency, no admin rights)</b></summary>

[Scoop](https://scoop.sh) installs into your user profile, needs no administrator prompt, and always puts commands on `PATH`. It is the only Windows package manager that carries all of SubMiner's optional dependencies, including `xz`, so it is the best choice if you want a single tool to manage everything.

```powershell
# One-time Scoop setup (skip if you already have it)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Invoke-RestMethod -Uri https://get.scoop.sh | Invoke-Expression

# mpv lives in the "extras" bucket; everything else is in "main"
scoop bucket add extras
scoop install extras/mpv main/ffmpeg

# Optional: yt-dlp for YouTube playback, xz for TsukiHime subtitle downloads
scoop install main/yt-dlp main/xz
```

Close and reopen your terminal, then verify with `mpv --version` and `ffmpeg -version`.

</details>

<details>
<summary><b>Manual install (download the zips yourself)</b></summary>

1. Download mpv from [mpv.io/installation](https://mpv.io/installation/) (the Windows builds link) and ffmpeg from [ffmpeg.org/download.html](https://ffmpeg.org/download.html).
2. Unzip each one somewhere permanent, for example `C:\Tools\mpv` and `C:\Tools\ffmpeg`. Note the folder that actually contains `mpv.exe` and the one containing `ffmpeg.exe` (for ffmpeg this is usually a `bin` subfolder).
3. Press `Win`, type **Edit the system environment variables**, and open it. Click **Environment Variables…**, select **Path** under **User variables**, click **Edit…**, then use **New** to add each of those two folders. Confirm with **OK** on every dialog. Microsoft documents this in more detail under [environment variables](https://learn.microsoft.com/windows/deployment/usmt/usmt-recognized-environment-variables).
4. Close and reopen your terminal, since `PATH` changes only apply to newly opened windows. Then check:

```powershell
mpv --version
ffmpeg -version
```

If you see `not recognized as the name of a cmdlet`, the folder you added is not the one holding the `.exe`. Reopen the Path editor and double-check.

::: tip mpv can skip PATH, ffmpeg cannot
If you would rather not edit `PATH` for mpv, set `mpv.executablePath` to the full path of `mpv.exe` during first-run setup instead.

There is no equivalent setting for ffmpeg: SubMiner invokes it by bare name when generating card audio and screenshots, so ffmpeg has to be on `PATH`. Without it, cards are still created but their audio and image fields come out empty. (`subsync.ffmpeg_path` only affects subtitle sync, not card media.)
:::

</details>

**Optional extras:** [MeCab for Windows](https://taku910.github.io/mecab/#download) with the UTF-8 dictionary improves annotation accuracy; it is not in any package manager, so install it from that page. `xz` is needed only for [TsukiHime](/tsukihime-integration) subtitle downloads and is not packaged by winget or Chocolatey, so use `scoop install main/xz` or download [XZ Utils](https://tukaani.org/xz/) and add its folder to `PATH`.

The `subminer` command-line launcher and its picker tools (`fzf`, `rofi`, `chafa`, `ffmpegthumbnailer`) are Linux/macOS only; on Windows you use the **SubMiner mpv** shortcut instead.

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

This checks for the app binary, mpv, ffmpeg, yt-dlp, fzf, rofi, your config file, and the mpv socket path. Only the app binary and mpv are hard failures; the rest are reported as optional. Fix any hard failures before continuing.

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

SubMiner verifies AppImage, launcher, and Linux support-asset downloads against `SHA256SUMS.txt`. On Linux those support assets include the launcher-managed runtime plugin copy under `SubMiner/plugin/subminer`, the rofi theme at `SubMiner/themes/subminer.rasi`, and the scoped Matroska thumbnailer registration under `SubMiner/thumbnailers`. If the binary is in a protected path, SubMiner shows the exact command to run rather than elevating itself.

The tray "Check for Updates" entry installs the new app automatically on Linux, macOS, and Windows. On Linux it replaces the running `.AppImage` in place via `electron-updater` and refreshes the managed support assets from `subminer-assets.tar.gz`; AppImages managed by a system package (for example the AUR `/opt/SubMiner/SubMiner.AppImage`) are skipped so the package manager stays in charge.

`subminer -u` also performs the AppImage, launcher, and managed support-asset updates directly from the launcher process, which is useful when SubMiner is not currently running.

## How It All Fits Together

SubMiner is an overlay that sits on top of mpv. It connects to mpv through an IPC socket, renders subtitles as interactive text using a bundled Yomitan dictionary engine, and optionally creates Anki flashcards via AnkiConnect.

The `subminer` launcher handles mpv IPC socket setup automatically. If you launch mpv yourself or from another tool, you must pass `--input-ipc-server=/tmp/subminer-socket` (or `\\.\pipe\subminer-socket` on Windows) - without it the overlay starts but subtitles won't appear.

The bundled mpv plugin is injected at runtime automatically - you don't need to install it separately. On Linux, the `subminer` launcher checks for its managed runtime plugin copy, rofi theme, and scoped thumbnailer registration before every mpv-managed launch and installs those support assets from the bundled app automatically if one is missing. For a rofi picker launch, this check runs before the picker opens. It provides in-player keybindings (the `y` chord) for controlling the overlay from within mpv. See [MPV Plugin](/mpv-plugin) for the full keybinding and configuration reference.

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

### Linux Support Assets

SubMiner ships the Linux rofi theme, scoped Matroska thumbnailer registration, and launcher-managed runtime plugin copy in `subminer-assets.tar.gz`:

```bash
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer-assets.tar.gz -O /tmp/subminer-assets.tar.gz
tar -xzf /tmp/subminer-assets.tar.gz -C /tmp
mkdir -p ~/.local/share/SubMiner/themes
cp /tmp/assets/themes/subminer.rasi ~/.local/share/SubMiner/themes/subminer.rasi
mkdir -p ~/.local/share/SubMiner/thumbnailers
cp /tmp/assets/thumbnailers/subminer-ffmpegthumbnailer.thumbnailer ~/.local/share/SubMiner/thumbnailers/
mkdir -p ~/.local/share/SubMiner/plugin
cp -R /tmp/plugin/subminer ~/.local/share/SubMiner/plugin/subminer
```

`subminer -u` and the tray updater keep those Linux support assets in sync automatically once the `SubMiner` data dir exists. Normal Linux launcher playback also auto-installs all three assets from the bundled app if one is missing, so manual extraction is mainly useful for pre-seeding or custom setups. Rofi receives the SubMiner data path through its process-local `XDG_DATA_DIRS`, so the thumbnailer registration does not change the desktop-wide configuration.

Override the theme path with `SUBMINER_ROFI_THEME=/absolute/path/to/theme.rasi`.

Next: [Usage](/usage) - learn about the `subminer` wrapper, keybindings, and YouTube playback.
