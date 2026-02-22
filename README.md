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

[![SubMiner demo (GIF preview)](./assets/minecard.gif)](./assets/minecard.mp4)

</div>

<br />

## What it does

SubMiner is an Electron overlay that sits on top of mpv. It turns your video player into a full sentence-mining workstation:

- **Hover to look up** — Yomitan dictionary popups directly on subtitles
- **One-key mining** — Creates Anki cards with sentence, audio, screenshot, and translation
- **N+1 highlighting** — Marks known words from your Anki deck so unknown ones jump out
- **Subtitle tools** — Download from Jimaku, sync with alass/ffsubsync, all in-player
- **Immersion tracking** — SQLite-powered stats on your watch time and mining activity
- **Texthooker page built in** — WebSocket streaming to external tools, no extra setup

## Quick start

### 1. Install

**Linux (AppImage):**

```bash
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/SubMiner-0.1.0.AppImage -O ~/.local/bin/SubMiner.AppImage
chmod +x ~/.local/bin/SubMiner.AppImage
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer -O ~/.local/bin/subminer
chmod +x ~/.local/bin/subminer
```

> [!NOTE]
> The `subminer` wrapper uses a [Bun](https://bun.sh) shebang. Make sure `bun` is on your `PATH`.

**From source** or **macOS** — see the [installation guide](https://docs.subminer.moe/installation#from-source).

### 2. Install the mpv plugin and configuration file

```bash
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer-assets-0.1.0.tar.gz -O /tmp/subminer-assets.tar.gz
tar -xzf /tmp/subminer-assets.tar.gz -C /tmp
cp /tmp/plugin/subminer.lua ~/.config/mpv/scripts/
cp /tmp/plugin/subminer.conf ~/.config/mpv/script-opts/
mkdir -p ~/.config/SubMiner && cp /tmp/config.example.jsonc ~/.config/SubMiner/config.jsonc
```

### 3. Set up Yomitan Dictionaries

```bash
subminer app --start --yomitan
```

### 4. Mine

```bash
subminer app --start --background
subminer video.mkv
```

## Requirements

| Required                                   | Optional                     |
| ------------------------------------------ | ---------------------------- |
| `mpv` with IPC socket                      | `yt-dlp`                     |
| `ffmpeg`                                   | `guessit` (better AniSkip title/episode detection) |
| `mecab` + `mecab-ipadic`                   | `fzf` / `rofi`               |
| Linux: `hyprctl` or `xdotool` + `xwininfo` | `chafa`, `ffmpegthumbnailer` |
| macOS: Accessibility permission            |                              |

## Documentation

For full guides on configuration, Anki, Jellyfin, and more, see [docs.subminer.moe](https://docs.subminer.moe).

## Acknowledgments

Built on the shoulders of [GameSentenceMiner](https://github.com/bpwhelan/GameSentenceMiner), [mpvacious](https://github.com/Ajatt-Tools/mpvacious), [Anacreon-Script](https://github.com/friedrich-de/Anacreon-Script), and [autosubsync-mpv](https://github.com/joaquintorres/autosubsync-mpv). Subtitles powered by [Jimaku.cc](https://jimaku.cc). Dictionary lookups via [Yomitan](https://github.com/yomidevs/yomitan).

## License

[GNU General Public License v3.0](LICENSE)
