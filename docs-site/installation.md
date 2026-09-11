# Installation

SubMiner draws an interactive overlay on top of the [mpv](https://mpv.io) video player. While you watch Japanese media, hover any word in the subtitles to look it up, then turn it into an Anki card without switching apps.

Building cards from the content you are actually watching is called **sentence mining**, and it is the whole point of SubMiner. It bundles its own copy of **Yomitan** (a pop-up dictionary) and talks to **AnkiConnect** (the add-on that lets other programs write cards into Anki), so the sentence, audio, and screenshot fields get filled in for you.

Three steps to get started:

1. **Install requirements** - mpv and a few optional extras
2. **Install SubMiner** - from the AUR, or download from GitHub Releases
3. **Launch the app** - first-run setup walks you through dictionaries, the launcher, and everything else

## 1. Install requirements

Only **mpv** is strictly required. Everything else is optional, though you will want ffmpeg unless you are fine with cards that have no audio or screenshot.

Some rows below matter only for the `subminer` command-line launcher's picker features. On Windows, the **SubMiner mpv** shortcut remains the recommended playback entry point.

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
Wayland has no universal API for window positioning. Each compositor exposes its own IPC, so SubMiner needs a backend per compositor. Only Hyprland and Sway have native Wayland backends. If you run a different Wayland compositor (GNOME, KDE Plasma, river, etc.), both mpv **and** SubMiner must run under X11 or Xwayland. The `subminer` launcher handles this automatically when `--backend x11` is set or the X11 backend is auto-detected.
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

The launcher's picker tools (`fzf`, `rofi`, `chafa`, `ffmpegthumbnailer`) are for Linux and macOS. On Windows, use the **SubMiner mpv** shortcut for playback or install the optional `subminer` terminal wrapper during setup.

## 2. Install SubMiner

### Arch Linux (AUR) {#arch-aur}

Install [`subminer-bin`](https://aur.archlinux.org/packages/subminer-bin) from the AUR. The package includes the SubMiner AppImage and its launcher wrapper. Bun is included with the app, so the package has no Bun dependency. Install updates through your AUR helper or package manager.

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
First-run setup can install the `subminer` command-line launcher for you. It uses Bun bundled with the AppImage, so it does not need a separate Bun installation or a Bun entry on `PATH`. The downloaded wrapper works the same way. See [manual launcher install](#manual-launcher-install-linux).
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
First-run setup can install the `subminer` command-line launcher for you. It uses Bun bundled inside `SubMiner.app`, so it does not need a separate Bun installation or a Bun entry on `PATH`. The downloaded wrapper works the same way. See [manual launcher install](#manual-launcher-install-macos).
:::

### Windows (installer) {#windows-installer}

Download the latest installer from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest):

- `SubMiner-<version>.exe` - installer (recommended)
- `SubMiner-<version>-win.zip` - portable fallback
- `subminer.cmd` - optional terminal launcher wrapper

Make sure `mpv.exe` is on your `PATH`, or set `mpv.executablePath` in the config during first-run setup.

### From source

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

Source and development commands use Bun installed on your system.

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

## 3. Launch and first-run setup

Launch SubMiner and the setup wizard opens on its own:

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
- **`subminer` launcher** _(optional)_ - installs a wrapper into a writable terminal PATH directory. The wrapper uses Bun packaged with the app, with no separate runtime setup. If the included runtime is unavailable, the launcher controls show an error asking you to reinstall SubMiner.
- **Windows shortcut** _(Windows only)_ - create a `SubMiner mpv` Start Menu/Desktop shortcut

The `Finish setup` button requires a config file and at least one Yomitan dictionary. The launcher is optional and never blocks setup completion.

On Linux and macOS, setup selects a writable directory already on your terminal `PATH`. If it cannot find one, it creates `~/.local/bin` and shows the `export PATH=...` command to run. Add that command to your shell configuration yourself if you want it in future terminals. Setup never edits shell configuration files. On Windows, setup adds only the wrapper directory to the user `PATH`. Setup stores a custom app location so the wrapper can find an AppImage or app bundle outside the usual install directories.

> [!TIP]
> You can re-open the setup wizard at any time with `subminer app --setup` or `SubMiner.AppImage --setup`.

### Play a video

Once setup is complete:

```bash
subminer video.mkv
```

The overlay appears over mpv. If a subtitle track loaded, its text shows up in the overlay as hoverable words.

On **Windows**, the recommended way to play video is with the **SubMiner mpv** shortcut created during setup - double-click it, or drag a video file onto it.

### Verify setup

Run the built-in diagnostic:

```bash
subminer doctor
```

This checks for the app binary, mpv, ffmpeg, yt-dlp, fzf, rofi, your config file, and the mpv socket path. Only the app binary and mpv are hard failures; the rest are reported as optional. Fix any hard failures before continuing.

## Anki setup (recommended)

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

The tray "Check for Updates" entry installs the new app automatically on Linux, macOS, and Windows. Current `subminer` wrappers remain small bootstraps that locate the installed app and its private runtime. On Linux the updater replaces the running `.AppImage` in place via `electron-updater` and refreshes managed support assets from `subminer-assets.tar.gz`. The next launcher invocation detects the changed AppImage fingerprint and prepares the matching Bun and CLI cache before running the command. App startup also refreshes this payload and migrates recognized writable legacy launchers. AppImages managed by a system package, for example the AUR `/opt/SubMiner/SubMiner.AppImage`, are skipped so the package manager stays in charge.

On Linux, `subminer -u` updates the AppImage and managed support assets directly, even when the app is not running. The launcher cache refreshes when the app fingerprint changes. AUR installs remain under package-manager control and should be updated through the package manager.

## How it all fits together

SubMiner is an overlay window that sits on top of mpv. It talks to mpv over an IPC socket, renders each subtitle line as interactive text backed by the bundled Yomitan dictionary engine, and writes Anki cards through AnkiConnect when you ask it to.

The `subminer` launcher handles mpv IPC socket setup automatically. If you launch mpv yourself or from another tool, you must pass `--input-ipc-server=/tmp/subminer-socket` (or `\\.\pipe\subminer-socket` on Windows) - without it the overlay starts but subtitles won't appear.

SubMiner injects the bundled mpv plugin at runtime, so there is nothing to install separately. On Linux, the `subminer` launcher checks for its managed runtime plugin copy, rofi theme, and scoped thumbnailer registration before every mpv-managed launch and installs those support assets from the bundled app automatically if one is missing. For a rofi picker launch, this check runs before the picker opens. The plugin adds in-player keybindings (the `y` chord) for driving the overlay from mpv. See [MPV Plugin](/mpv-plugin) for the full keybinding and configuration reference.

## Platform notes

### macOS

**MeCab paths (Homebrew):**

- Apple Silicon (M1/M2): `/opt/homebrew/bin/mecab`
- Intel: `/usr/local/bin/mecab`

`mecab` has to be on your PATH when SubMiner launches.

**Fullscreen:** The overlay follows mpv into fullscreen. If it does not, accessibility permission is the usual cause.

### Windows

- The **SubMiner mpv** shortcut is the recommended way to launch playback. It starts `mpv.exe` with the right IPC socket and subtitle defaults.
- First-run setup adds only `%LOCALAPPDATA%\SubMiner\bin` to the HKCU user PATH. It does not add `SubMiner.exe` to PATH.
- IPC socket on Windows is `\\.\pipe\subminer-socket` - do not use `/tmp/subminer-socket`.
- Config is stored at `%APPDATA%\SubMiner\config.jsonc`.

## Manual launcher install

Current launcher downloads use Bun included in the SubMiner app. The wrapper searches normal install locations and honors `SUBMINER_BINARY_PATH`; Linux also honors `SUBMINER_APPIMAGE_PATH`.

### Linux {#manual-launcher-install-linux}

```bash
# Download the launcher
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer -O ~/.local/bin/subminer
chmod +x ~/.local/bin/subminer
```

### macOS {#manual-launcher-install-macos}

```bash
# Download the launcher
sudo curl -fSL https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer -o /usr/local/bin/subminer
sudo chmod +x /usr/local/bin/subminer
```

### Windows {#manual-launcher-install-windows}

Download `subminer.cmd` from GitHub Releases and place it in a directory on your user `PATH`. It finds the installed app in the normal per-user or Program Files location. Set `SUBMINER_BINARY_PATH` if you use a portable or custom install.

Launchers installed before the private-runtime change may still be bundled JavaScript with a Bun shebang. Those old files need system Bun until a current app startup migrates a recognized writable launcher, or until you replace one with the current release wrapper.

## Optional extras

### Linux support assets

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
