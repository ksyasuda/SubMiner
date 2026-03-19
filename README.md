<div align="center">
  <img src="assets/SubMiner.png" width="140" alt="SubMiner logo">

  # SubMiner

  **Sentence-mine from mpv — look up words, one-key Anki export, immersion tracking.**

  [![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
  [![Linux](https://img.shields.io/badge/platform-Linux%20%7C%20macOS%20%7C%20Windows-informational)]()
  [![Docs](https://img.shields.io/badge/docs-docs.subminer.moe-blueviolet)](https://docs.subminer.moe)
  [![AUR](https://img.shields.io/aur/version/subminer-bin)](https://aur.archlinux.org/packages/subminer-bin)

</div>

---

SubMiner is an Electron overlay for [mpv](https://mpv.io) that turns video into a sentence-mining workstation. Look up any word with [Yomitan](https://github.com/yomidevs/yomitan), mine it to Anki with one key, and track your immersion over time.

<div align="center">

[![SubMiner demo (Animated preview)](./assets/minecard.webp)](./assets/minecard.mp4)

</div>

## Features

**Dictionary lookups** — Yomitan runs inside the overlay. Hover or navigate to any word for full dictionary popups without leaving mpv.

**One-key Anki mining** — Press one key to create a card with the sentence, audio clip, screenshot, and machine translation from the exact playback moment.

<div align="center">
  <img src="docs-site/public/screenshots/yomitan-lookup.png" width="800" alt="Yomitan popup with dictionary entry and mine button over annotated subtitles in mpv">
</div>

**Reading annotations** — Real-time subtitle annotations with N+1 targeting, frequency highlighting, JLPT tags, and a character name dictionary. Grammar-only tokens render as plain text.

<div align="center">
  <img src="docs-site/public/screenshots/annotations.png" width="800" alt="Annotated subtitles with frequency highlighting, JLPT underlines, known words, and N+1 targets">
</div>

**Immersion dashboard** — Local stats dashboard with watch time, anime progress, vocabulary growth, mining throughput, and session history.

<div align="center">
  <img src="docs-site/public/screenshots/stats-overview.png" width="800" alt="Stats dashboard with watch time, cards mined, streaks, and tracking snapshot">
</div>

**Integrations** — AniList episode tracking, Jellyfin remote playback, Jimaku subtitle downloads, alass/ffsubsync, and an annotated websocket feed for external clients.

<div align="center">
  <img src="docs-site/public/screenshots/texthooker.png" width="800" alt="Texthooker page with annotated subtitle lines and frequency highlighting">
</div>

---

## Quick Start

### Install

<details>
<summary><b>Arch Linux (AUR)</b></summary>

```bash
paru -S subminer-bin
```

Or manually:

```bash
git clone https://aur.archlinux.org/subminer-bin.git && cd subminer-bin && makepkg -si
```

</details>

<details>
<summary><b>Linux (AppImage)</b></summary>

```bash
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/SubMiner.AppImage -O ~/.local/bin/SubMiner.AppImage \
	&& chmod +x ~/.local/bin/SubMiner.AppImage
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer -O ~/.local/bin/subminer \
	&& chmod +x ~/.local/bin/subminer
```

> [!NOTE]
> The `subminer` wrapper uses a [Bun](https://bun.sh) shebang. Make sure `bun` is on your `PATH`.

</details>

<details>
<summary><b>macOS / Windows / From source</b></summary>

**macOS** — Download the latest DMG/ZIP from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest) and drag `SubMiner.app` into `/Applications`.

**Windows** — Download the latest installer or portable `.zip` from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest). Keep `mpv` on `PATH`.

**From source** — See [docs.subminer.moe/installation#from-source](https://docs.subminer.moe/installation#from-source).

</details>

### First Launch

Run `SubMiner.AppImage` (Linux), `SubMiner.app` (macOS), or `SubMiner.exe` (Windows). On first launch, SubMiner starts in the tray, creates a default config, and opens a setup popup where you can install the mpv plugin and configure Yomitan dictionaries.

### Mine

```bash
subminer video.mkv          # auto-starts overlay + resumes playback
subminer --start video.mkv  # explicit overlay start (if plugin auto_start=no)
subminer stats              # open the immersion dashboard
```

---

## Requirements

| Required | Optional |
|---|---|
| [`mpv`](https://mpv.io) with IPC socket | `yt-dlp` |
| `ffmpeg` | `guessit` (AniSkip detection) |
| `mecab` + `mecab-ipadic` | `fzf` / `rofi` |
| [`bun`](https://bun.sh) (source builds, Linux wrapper) | `chafa`, `ffmpegthumbnailer` |
| Linux: `hyprctl` or `xdotool` + `xwininfo` | |
| macOS: Accessibility permission | |

Windows uses native window tracking and does not need the Linux compositor tools.

## Documentation

Full guides on configuration, Anki, Jellyfin, immersion tracking, and more at **[docs.subminer.moe](https://docs.subminer.moe)**.

## Acknowledgments

Built on [GameSentenceMiner](https://github.com/bpwhelan/GameSentenceMiner), [Renji's Texthooker Page](https://github.com/Renji-XD/texthooker-ui), [Anacreon-Script](https://github.com/friedrich-de/Anacreon-Script), and [Bee's Character Dictionary](https://github.com/bee-san/Japanese_Character_Name_Dictionary). Subtitles from [Jimaku.cc](https://jimaku.cc). Lookups via [Yomitan](https://github.com/yomidevs/yomitan). JLPT tags from [yomitan-jlpt-vocab](https://github.com/stephenmk/yomitan-jlpt-vocab).

## License

[GNU General Public License v3.0](LICENSE)
