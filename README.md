<div align="center">
  <img src="assets/SubMiner.png" width="169" alt="SubMiner logo">
  <h1>SubMiner</h1>
  <strong>Look up words, mine to Anki, and enrich cards with context — without leaving mpv.</strong>
  <br /><br />

[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
[![Linux](https://img.shields.io/badge/platform-Linux%20%7C%20macOS-informational)]()
[![Docs](https://img.shields.io/badge/docs-docs.subminer.moe-blueviolet)](https://docs.subminer.moe)

</div>

<br />

<div align="center">

[![SubMiner demo (Animated preview)](./assets/minecard.webp)](./assets/minecard.mp4)

</div>

<br />

## What it does

SubMiner is an Electron overlay that sits on top of mpv. It turns your video player into a full sentence-mining workstation:

- **Dictionary lookups** — Yomitan popups on subtitles with hover or full keyboard-driven navigation; hover-aware auto-pause keeps playback in sync
- **One-key mining** — Creates Anki cards with sentence, audio, screenshot, and AI-powered translation
- **Reading annotations** — N+1 targeting, frequency highlighting, and JLPT underlining while you watch
- **Subtitle tools** — Jimaku downloads, alass/ffsubsync sync, and YouTube subtitle generation via manual-track reuse plus whisper.cpp fallback with optional AI cleanup
- **Texthooker** — Built-in texthooker page and annotated websocket API for external clients
- **Immersion tracking** — SQLite-powered stats on watch time and mining activity
- **Integrations** — Jellyfin remote playback, AniList episode progress, and AnkiConnect auto-enrichment

## Quick start

### 1. Install

**Linux (AppImage):**

```bash
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/SubMiner.AppImage -O ~/.local/bin/SubMiner.AppImage
chmod +x ~/.local/bin/SubMiner.AppImage
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer -O ~/.local/bin/subminer
chmod +x ~/.local/bin/subminer

```

> [!NOTE]
> The `subminer` wrapper uses a [Bun](https://bun.sh) shebang. Make sure `bun` is on your `PATH`.

**From source** or **macOS** — initialize submodules first (`git submodule update --init --recursive`). Bundled Yomitan is built natively with Bun from the `vendor/subminer-yomitan` submodule into `build/yomitan` during `bun run build`, so Bun is the only JS runtime/package manager required for source builds. Full install guide: [docs.subminer.moe/installation#from-source](https://docs.subminer.moe/installation#from-source).

### 2. Launch the app once

```bash
SubMiner.AppImage
```

On first launch, SubMiner now:

- starts in the tray/background
- creates the default config directory and `config.jsonc`
- opens a compact setup popup
- can install the mpv plugin to the default mpv scripts location for you
- links directly to Yomitan settings so you can install dictionaries before finishing setup

Existing installs that already have a valid config plus at least one Yomitan dictionary are auto-detected as complete and will not be re-prompted.

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
```

## Requirements

| Required                                   | Optional                                           |
| ------------------------------------------ | -------------------------------------------------- |
| `bun`                                      |                                                    |
| `mpv` with IPC socket                      | `yt-dlp`                                           |
| `ffmpeg`                                   | `guessit` (better AniSkip title/episode detection) |
| `mecab` + `mecab-ipadic`                   | `fzf` / `rofi`                                     |
| Linux: `hyprctl` or `xdotool` + `xwininfo` | `chafa`, `ffmpegthumbnailer`                       |
| macOS: Accessibility permission            |                                                    |

## Documentation

For full guides on configuration, Anki, Jellyfin, and more, see [docs.subminer.moe](https://docs.subminer.moe). Contributor setup, build, and testing docs now live in the docs repo: [docs.subminer.moe/development#testing](https://docs.subminer.moe/development#testing).

## Acknowledgments

Built on the shoulders of [GameSentenceMiner](https://github.com/bpwhelan/GameSentenceMiner), [Renji's Texthooker Page](https://github.com/Renji-XD/texthooker-ui), [mpvacious](https://github.com/Ajatt-Tools/mpvacious), [Anacreon-Script](https://github.com/friedrich-de/Anacreon-Script), and [Bee's Character Dictionary](https://github.com/bee-san/Japanese_Character_Name_Dictionary). Subtitles powered by [Jimaku.cc](https://jimaku.cc). Dictionary lookups via [Yomitan](https://github.com/yomidevs/yomitan).

## License

[GNU General Public License v3.0](LICENSE)
