# Installation

SubMiner draws an interactive overlay on top of the [mpv](https://mpv.io) video player. While you watch Japanese media, you hover a word in the subtitles to look it up, then turn it into an Anki card without leaving the video.

Building cards from what you watch is called **sentence mining**. SubMiner bundles its own copy of **Yomitan** (a pop-up dictionary) and talks to **AnkiConnect** (an Anki add-on that lets other programs create cards), so it can fill in the sentence, audio, and screenshot for you.

Getting started takes three steps:

1. Install mpv and the optional extras you want.
2. Install SubMiner.
3. Launch it and follow the first-run setup.

## 1. Install requirements

Only mpv is required. Install ffmpeg too unless you are fine with cards that have no audio or screenshot.

| Dependency               | Needed for                                                                                  | Platforms    |
| ------------------------ | ------------------------------------------------------------------------------------------- | ------------ |
| mpv                      | Required. The player SubMiner draws over.                                                   | All          |
| fuse2                    | Required to run the AppImage.                                                               | Linux        |
| ffmpeg                   | Recommended. Audio clips and screenshots on cards. Without it those fields stay empty.      | All          |
| MeCab + mecab-ipadic     | Recommended. More accurate N+1, JLPT, and frequency highlighting.                           | All          |
| yt-dlp                   | YouTube playback.                                                                           | All          |
| xz                       | [TsukiHime](/tsukihime-integration) subtitle downloads. Most Linux distros already have it. | All          |
| guessit                  | Better title, season, and episode detection for [AniSkip](/aniskip-integration).            | All          |
| alass or ffsubsync       | Subtitle syncing. You need at least one to use it.                                          | All          |
| fzf, rofi                | The file pickers in the `subminer` command (rofi is Linux only).                            | Linux, macOS |
| chafa, ffmpegthumbnailer | Thumbnail previews in the pickers.                                                          | Linux, macOS |

To generate Japanese subtitles from audio, you also need whisper.cpp. See [Subtitle generation](/subtitle-generation).

### Linux

SubMiner needs to track the mpv window, and how it does that depends on your desktop:

- **Hyprland**: supported natively through `hyprctl`.
- **Sway**: supported natively through `swaymsg`.
- **Anything else** (X11, GNOME, KDE Plasma, other Wayland compositors): mpv and SubMiner must run under X11 or Xwayland. Install `xdotool` and `xwininfo`. The `subminer` command picks the X11 backend automatically, or you can force it with `--backend x11`.

<details>
<summary><b>Arch Linux</b></summary>

```bash
sudo pacman -S --needed mpv ffmpeg
# Recommended
sudo pacman -S --needed mecab mecab-ipadic
# Optional
sudo pacman -S --needed yt-dlp fzf rofi chafa ffmpegthumbnailer
# Optional: subtitle sync (install at least one)
paru -S --needed alass python-ffsubsync
# Only for desktops other than Hyprland or Sway
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
# Only for desktops other than Hyprland or Sway
sudo apt install xdotool x11-utils
# Optional: subtitle sync
pip install ffsubsync
cargo install alass-cli
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
# Only for desktops other than Hyprland or Sway
sudo dnf install xdotool xorg-x11-utils
# Optional: subtitle sync
pip install ffsubsync
cargo install alass-cli
```

</details>

### macOS

You need macOS 11 (Big Sur) or later.

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

`mecab` must be on your `PATH` when SubMiner starts. Homebrew puts it in `/opt/homebrew/bin` on Apple Silicon and `/usr/local/bin` on Intel.

### Windows

You need Windows 10 or later. Install mpv and ffmpeg with [winget](https://learn.microsoft.com/windows/package-manager/winget/), which ships with Windows 11 and current Windows 10. In PowerShell or Command Prompt:

```powershell
winget install shinchiro.mpv
winget install Gyan.FFmpeg
winget install yt-dlp.yt-dlp   # optional, for YouTube
```

Close and reopen the terminal, then check both commands work:

```powershell
mpv --version
ffmpeg -version
```

ffmpeg must be on `PATH`, because SubMiner runs it by name to make card audio and screenshots. mpv does not have to be. If `mpv --version` says `not recognized`, find `mpv.exe` (usually in `%LOCALAPPDATA%\Programs\mpv`) and either add that folder to `PATH` or enter the full path to `mpv.exe` during first-run setup (`mpv.executablePath`).

<details>
<summary><b>Alternative: Scoop (no admin rights, includes xz)</b></summary>

[Scoop](https://scoop.sh) installs into your user profile and always adds commands to `PATH`. It is the only Windows package manager that also packages `xz`, which [TsukiHime](/tsukihime-integration) downloads need.

```powershell
# One-time Scoop setup
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Invoke-RestMethod -Uri https://get.scoop.sh | Invoke-Expression

scoop bucket add extras
scoop install extras/mpv main/ffmpeg
# Optional
scoop install main/yt-dlp main/xz
```

</details>

<details>
<summary><b>Alternative: manual download</b></summary>

1. Download mpv from [mpv.io/installation](https://mpv.io/installation/) and ffmpeg from [ffmpeg.org/download.html](https://ffmpeg.org/download.html).
2. Unzip each into a permanent folder, for example `C:\Tools\mpv` and `C:\Tools\ffmpeg`. Find the folders that contain `mpv.exe` and `ffmpeg.exe` (for ffmpeg this is usually `bin`).
3. Press `Win`, search for **Edit the system environment variables**, and open it. Click **Environment Variables**, select **Path** under **User variables**, click **Edit**, and add both folders with **New**.
4. Open a new terminal and run `mpv --version` and `ffmpeg -version`. If either says `not recognized`, the folder you added does not contain the `.exe`.

For `xz` without Scoop, download [XZ Utils](https://tukaani.org/xz/) and add its folder to `PATH` the same way.

</details>

For more accurate highlighting, install [MeCab for Windows](https://taku910.github.io/mecab/#download) with the UTF-8 dictionary. The fzf and rofi pickers do not apply on Windows.

## 2. Install SubMiner

### Arch Linux (AUR) {#arch-aur}

Install [`subminer-bin`](https://aur.archlinux.org/packages/subminer-bin). It includes the AppImage and the `subminer` command.

```bash
paru -S subminer-bin
```

### Linux (AppImage) {#linux-appimage}

```bash
mkdir -p ~/.local/bin
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/SubMiner.AppImage -O ~/.local/bin/SubMiner.AppImage
chmod +x ~/.local/bin/SubMiner.AppImage
```

First-run setup can install the `subminer` command for you.

### macOS (DMG) {#macos-dmg}

1. Download the DMG from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest), open it, and drag `SubMiner.app` into `/Applications`.
2. If macOS blocks the app on first launch, right-click it and choose **Open**, or run:

   ```bash
   xattr -d com.apple.quarantine /Applications/SubMiner.app
   ```

3. Open **System Settings > Privacy & Security > Accessibility** and enable SubMiner (add it if it is missing). The overlay cannot follow the mpv window without this.

First-run setup can install the `subminer` command for you.

### Windows (installer) {#windows-installer}

Download from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest):

- `SubMiner-<version>.exe`: the installer. Use this one.
- `SubMiner-<version>-win.zip`: portable version.
- `subminer.cmd`: optional terminal command (setup can install it for you).

### From source

<details>
<summary><b>Linux</b></summary>

```bash
git clone --recurse-submodules https://github.com/ksyasuda/SubMiner.git
cd SubMiner
make deps
bun run build
bun run build:appimage   # optional: package an AppImage
```

Building from source needs [Bun](https://bun.sh) installed.

</details>

<details>
<summary><b>macOS</b></summary>

```bash
git clone --recurse-submodules https://github.com/ksyasuda/SubMiner.git
cd SubMiner
make deps
make build-macos
```

The `.dmg` and `.zip` land in `release/`. For an unsigned local build, run `bun run build:mac:unsigned`.

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

## 3. Launch and first-run setup

Start SubMiner. The setup window opens on first launch.

- **Linux (AUR)**: `subminer app --setup`
- **Linux (AppImage)**: `~/.local/bin/SubMiner.AppImage --setup`
- **macOS**: open `SubMiner.app` from `/Applications`
- **Windows**: run SubMiner from the Start menu

Setup walks you through:

1. **Config file.** Created at `~/.config/SubMiner/config.jsonc` (Linux and macOS) or `%APPDATA%\SubMiner\config.jsonc` (Windows).
2. **Yomitan dictionaries.** Import at least one dictionary, or lookups will not work. SubMiner's Yomitan is separate from any Yomitan in your browser.
3. **The `subminer` command** (optional). Setup installs it into a folder already on your `PATH`. If there is none on Linux or macOS, it uses `~/.local/bin` and shows the `export PATH=...` line to add to your shell config. On Windows it adds `%LOCALAPPDATA%\SubMiner\bin` to your user `PATH`.
4. **SubMiner mpv shortcut** (Windows only). A Start menu or desktop shortcut that opens mpv with SubMiner attached.

**Finish setup** unlocks once the config exists and at least one dictionary is imported. To reopen setup later, run `subminer app --setup`.

### Play a video

```bash
subminer video.mkv
```

On Windows, double-click the **SubMiner mpv** shortcut or drag a video onto it.

The overlay appears over mpv, and the subtitle text becomes hoverable. See [Usage](/usage) for everyday use.

### Check your setup

```bash
subminer doctor
```

This checks for the SubMiner app, mpv, ffmpeg, yt-dlp, fzf, rofi, your config file, and the mpv socket path. Only a missing app or mpv counts as a failure. The rest are reported as optional.

## Anki setup

To create cards:

1. Install [Anki](https://apps.ankiweb.net/).
2. In Anki, open **Tools > Add-ons > Get Add-ons** and enter `2055492159` to install [AnkiConnect](https://ankiweb.net/shared/info/2055492159).
3. Restart Anki. Keep it open while you use SubMiner.

SubMiner connects to AnkiConnect at its default address with no extra setup. To choose your deck and card fields, see [Anki integration](/anki-integration).

## Updates

```bash
subminer -u
```

The tray menu's **Check for Updates** also installs updates on Linux, macOS, and Windows. If the AppImage sits in a folder you cannot write to, SubMiner prints the command to run instead of asking for admin rights.

If you installed from the AUR, update through your package manager instead.

## Launching mpv yourself

The `subminer` command and the Windows shortcut start mpv with the IPC socket SubMiner needs. If you start mpv another way, add this option or the overlay starts without subtitles:

```bash
--input-ipc-server=/tmp/subminer-socket      # Linux and macOS
--input-ipc-server=\\.\pipe\subminer-socket  # Windows
```

SubMiner loads its mpv plugin automatically, so there is nothing else to install. See [mpv plugin](/mpv-plugin) for the in-player keybindings.

## Manual launcher install

Use these if you skipped the launcher during setup. The launcher finds SubMiner in the usual install locations. For a custom location, set `SUBMINER_BINARY_PATH` to the app executable.

### Linux {#manual-launcher-install-linux}

```bash
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer -O ~/.local/bin/subminer
chmod +x ~/.local/bin/subminer
```

### macOS {#manual-launcher-install-macos}

```bash
sudo curl -fSL https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer -o /usr/local/bin/subminer
sudo chmod +x /usr/local/bin/subminer
```

### Windows {#manual-launcher-install-windows}

Download `subminer.cmd` from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest) and put it in a folder on your user `PATH`.

## Bundled Bun runtime {#bundled-bun-runtime}

The `subminer` command runs on a copy of [Bun](https://bun.sh) 1.3.5 that ships inside the app, so you do not need to install Bun. Bun is MIT licensed and statically links JavaScriptCore (LGPL 2.0) and TinyCC (LGPL 2.1). License texts and a `SOURCE.md` ship in the app under `resources/bun/licenses`, and each GitHub release includes `bun-v1.3.5-source.tar.gz` with the matching sources.

Next: [Usage](/usage).
