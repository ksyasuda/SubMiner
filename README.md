<div align="center">
  <img src="assets/SubMiner.png" width="169" alt="SubMiner logo">
  <h1>SubMiner</h1>
  <strong>Look up words, mine to Anki, and enrich cards with context — without leaving mpv.</strong>
  <br /><br />

[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
[![Linux](https://img.shields.io/badge/platform-Linux%20%7C%20macOS%20%7C%20Windows-informational)]()
[![Docs](https://img.shields.io/badge/docs-docs.subminer.moe-blueviolet)](https://docs.subminer.moe)

</div>

<br />

<div align="center">

[![SubMiner demo (Animated preview)](./assets/minecard.webp)](./assets/minecard.mp4)

</div>

<br />

## What it does

SubMiner is an Electron overlay that sits on top of mpv. It turns your video player into a full sentence-mining workstation:

- **Look up words as you watch** — Yomitan dictionary popups on hover or keyboard-driven token-by-token navigation
- **One-key Anki mining** — Creates cards with sentence, audio, screenshot, and translation; optional local AnkiConnect proxy auto-enriches Yomitan cards instantly
- **Reading annotations** — N+1 targeting, frequency-dictionary highlighting, JLPT underlining, and character name dictionary for anime/manga proper nouns
- **Immersion stats** — Optional local dashboard and overlay for watch time, anime progress, session drill-down, vocabulary growth, and mining throughput
- **Subtitle tools** — Download from Jimaku, sync with alass/ffsubsync
- **Jellyfin & AniList integration** — Remote playback, cast device mode, and automatic episode progress tracking
- **Texthooker & API** — Built-in texthooker page and annotated websocket feed for external clients

## Quick start

### 1. Install

**Arch Linux (AUR):**

Install [`subminer-bin`](https://aur.archlinux.org/packages/subminer-bin) from the AUR. It installs the packaged AppImage plus the `subminer` wrapper:

```bash
paru -S subminer-bin
```

Or manually:

```bash
git clone https://aur.archlinux.org/subminer-bin.git
cd subminer-bin
makepkg -si
```

**Linux (AppImage):**

```bash
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/SubMiner.AppImage -O ~/.local/bin/SubMiner.AppImage
chmod +x ~/.local/bin/SubMiner.AppImage
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer -O ~/.local/bin/subminer
chmod +x ~/.local/bin/subminer

```

> [!NOTE]
> The `subminer` wrapper uses a [Bun](https://bun.sh) shebang. Make sure `bun` is on your `PATH`.

**macOS (DMG/ZIP):** download the latest packaged build from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest) and drag `SubMiner.app` into `/Applications`.

**Windows (Installer/ZIP):** download the latest `SubMiner-<version>.exe` installer or portable `.zip` from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest). Keep `mpv` installed and available on `PATH`.

**From source** — initialize submodules first (`git submodule update --init --recursive`). Bundled Yomitan is built from the `vendor/subminer-yomitan` submodule into `build/yomitan` during `bun run build`, so source builds only need Bun for the JS toolchain. Packaged macOS and Windows installs do not require Bun. Windows installer builds go through `electron-builder`; its bundled `app-builder-lib` NSIS templates already use the third-party `WinShell` plugin for shortcut AppUserModelID assignment, and the `WinShell.dll` binary is supplied by electron-builder's cached `nsis-resources` bundle, so `bun run build:win` does not need a separate repo-local plugin install step. Full install guide: [docs.subminer.moe/installation#from-source](https://docs.subminer.moe/installation#from-source).

### 2. Launch the app once

```bash
# Linux
SubMiner.AppImage
```

On macOS, launch `SubMiner.app`. On Windows, launch `SubMiner.exe` from the Start menu or install directory.

On first launch, SubMiner:

- starts in the tray/background
- creates the default config directory and `config.jsonc`
- opens a compact setup popup
- can install the mpv plugin to the default mpv scripts location for you
- links directly to Yomitan settings so you can install dictionaries before finishing setup

### 3. Finish setup

- click `Install mpv plugin` if you want the default plugin auto-start flow
- click `Open Yomitan Settings` and install at least one dictionary
- click `Refresh status`
- click `Finish setup`

The mpv plugin step is optional. Yomitan must report at least one installed dictionary before setup can be completed.

### 4. Mine

```bash
subminer video.mkv # default plugin config auto-starts visible overlay + resumes playback when ready
subminer --start video.mkv # optional explicit overlay start when plugin auto_start=no
subminer stats # open the local stats dashboard in your browser
```

## Requirements

| Required                                   | Optional                                           |
| ------------------------------------------ | -------------------------------------------------- |
| `bun` (source builds, Linux `subminer`)    |                                                    |
| `mpv` with IPC socket                      | `yt-dlp`                                           |
| `ffmpeg`                                   | `guessit` (better AniSkip title/episode detection) |
| `mecab` + `mecab-ipadic`                   | `fzf` / `rofi`                                     |
| Linux: `hyprctl` or `xdotool` + `xwininfo` | `chafa`, `ffmpegthumbnailer`                       |
| macOS: Accessibility permission            |                                                    |

Windows builds use native window tracking and do not require the Linux compositor helper tools.

## Documentation

For full guides on configuration, Anki, Jellyfin, immersion tracking/stats, and more, see [docs.subminer.moe](https://docs.subminer.moe). The VitePress source for that site lives in [`docs-site/`](./docs-site/).

## Acknowledgments

Built on the shoulders of [GameSentenceMiner](https://github.com/bpwhelan/GameSentenceMiner), [Renji's Texthooker Page](https://github.com/Renji-XD/texthooker-ui), [Anacreon-Script](https://github.com/friedrich-de/Anacreon-Script), and [Bee's Character Dictionary](https://github.com/bee-san/Japanese_Character_Name_Dictionary). Subtitles powered by [Jimaku.cc](https://jimaku.cc). Dictionary lookups via [Yomitan](https://github.com/yomidevs/yomitan), and JLPT tags from [yomitan-jlpt-vocab](https://github.com/stephenmk/yomitan-jlpt-vocab).

## License

[GNU General Public License v3.0](LICENSE)
