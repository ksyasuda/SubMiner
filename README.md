<div align="center">

<img src="assets/SubMiner.png" width="160" alt="SubMiner logo">

# SubMiner

Look up words with Yomitan, export to Anki in one key, track your immersion — all without leaving mpv.

[Installation](#quick-start) · [Requirements](#requirements) · [Usage](https://docs.subminer.moe/usage) · [Documentation](https://docs.subminer.moe)

[![Downloads](https://img.shields.io/github/downloads/ksyasuda/SubMiner/total?style=flat-square&color=1a1a2e)](https://github.com/ksyasuda/SubMiner/releases)
[![Release](https://img.shields.io/github/v/release/ksyasuda/SubMiner?style=flat-square&color=1a1a2e)](https://github.com/ksyasuda/SubMiner/releases/latest)
[![AUR](https://img.shields.io/aur/version/subminer-bin?style=flat-square&color=1a1a2e)](https://aur.archlinux.org/packages/subminer-bin)
[![Platform](https://img.shields.io/badge/platform-Linux%20·%20macOS%20·%20Windows-1a1a2e?style=flat-square)](https://github.com/ksyasuda/SubMiner)
[![License](https://img.shields.io/github/license/ksyasuda/SubMiner?style=flat-square&color=1a1a2e)](https://www.gnu.org/licenses/gpl-3.0)
[![TypeScript](https://img.shields.io/badge/TypeScript-1a1a2e?style=flat-square&logo=typescript&logoColor=3178c6)](https://www.typescriptlang.org)

[![SubMiner demo](./assets/minecard.webp)](https://github.com/user-attachments/assets/89e61895-e2b7-4b47-8d50-a35afe4132b2)

</div>

## Features

### Dictionary Lookups

Yomitan runs inside the overlay. Trigger a lookup on any word for full dictionary popups — definitions, pitch accent, frequency data — without ever leaving mpv.

<div align="center">
  <img src="docs-site/public/screenshots/yomitan-lookup.png" width="800" alt="Yomitan dictionary popup over annotated subtitles in mpv">
</div>

<br>

### Instant Anki Mining

Create an Anki card with the sentence, audio clip, screenshot, and machine translation from the exact playback moment with one key press, click, or controller input.

<div align="center">
  <img src="docs-site/public/screenshots/one-key-mining.png" width="800" alt="Anki card created from SubMiner with sentence, audio, and screenshot">
</div>

<br>

### Reading Annotations

Real-time subtitle annotations with frequency highlighting, JLPT tags, N+1 targeting, and a character name dictionary. Known words fade back; new words stand out. Grammar-only tokens render as plain text so you focus on what matters.

<div align="center">
  <img src="docs-site/public/screenshots/annotations.png" width="800" alt="Annotated subtitles with frequency coloring, JLPT underlines, and N+1 targets">
</div>

<br>

### Immersion Dashboard

Local stats dashboard — watch time, anime library, vocabulary growth, mining throughput, session history, and trends. All stored locally, no third-party tracking.

<div align="center">
  <img src="docs-site/public/screenshots/stats-overview.png" width="800" alt="Stats dashboard showing watch time, cards mined, streaks, and tracking data">
</div>

<br>

### Playlist Browser

Browse sibling episode files and the active mpv queue in one overlay modal. Open it with `Ctrl+Alt+P` to append episodes from the current directory, jump to queued items, remove entries, or reorder the playlist without leaving playback.

<br>

### Integrations

<table>
  <tr>
    <td><b>YouTube</b></td>
    <td>Auto-loaded yt-dlp subtitle tracks at startup with config-driven primary/secondary language priorities and a manual overlay picker on demand (<code>Ctrl+Alt+C</code>)</td>
  </tr>
  <tr>
    <td><b>AniList</b></td>
    <td>Automatic episode tracking and progress sync</td>
  </tr>
  <tr>
    <td><b>Jellyfin</b></td>
    <td>Browse and launch media from your Jellyfin server</td>
  </tr>
  <tr>
    <td><b>Jimaku</b></td>
    <td>Search and download Japanese subtitles</td>
  </tr>
  <tr>
    <td><b>alass / ffsubsync</b></td>
    <td>Automatic subtitle retiming — requires <code>alass</code> or <code>ffsubsync</code> on your <code>PATH</code> (optional; subtitle syncing is disabled without them)</td>
  </tr>
  <tr>
    <td><b>WebSocket</b></td>
    <td>Annotated subtitle feed for external clients (texthooker pages, custom tools)</td>
  </tr>
</table>

<div align="center">
  <img src="docs-site/public/screenshots/texthooker.png" width="800" alt="Texthooker page receiving annotated subtitle lines via WebSocket">
</div>

<br>

---

## Requirements

|                | Required                                | Recommended                              | Optional                                                    |
| -------------- | --------------------------------------- | ---------------------------------------- | ----------------------------------------------------------- |
| **Player**     | [`mpv`](https://mpv.io) with IPC socket | —                                        | —                                                           |
| **Processing** | —                                       | `ffmpeg` (audio clips & screenshots)     | `mecab` + `mecab-ipadic` (annotation POS filtering), `guessit` (AniSkip), `alass` / `ffsubsync` (subtitle sync) |
| **Media**      | —                                       | —                                        | `yt-dlp`, `chafa`, `ffmpegthumbnailer`                      |
| **Selection**  | —                                       | —                                        | `fzf` / `rofi`                                              |

> [!TIP]
> `ffmpeg` is not strictly required to run SubMiner, but without it audio clips and screenshots will not be attached to Anki cards. Most users will want it installed.

> [!NOTE]
> [`bun`](https://bun.sh) is required if building from source or using the CLI wrapper: `subminer`. Pre-built releases (AppImage, DMG, installer) do not require it.

**Platform-specific:**

| Linux                               | macOS                    | Windows       |
| ----------------------------------- | ------------------------ | ------------- |
| `hyprctl` or `xdotool` + `xwininfo` | Accessibility permission | No extra deps |

<details>
<summary><b>Arch Linux</b></summary>

```bash
paru -S --needed mpv ffmpeg
# Optional
paru -S --needed mecab-git mecab-ipadic yt-dlp fzf rofi chafa ffmpegthumbnailer xdotool xorg-xwininfo
# Optional: subtitle sync (install at least one for subtitle syncing to work)
paru -S --needed alass python-ffsubsync
# X11 / XWAYLAND
paru -S --needed xdotool xorg-xwininfo
```

</details>

<details>
<summary><b>macOS</b></summary>

```bash
brew install mpv ffmpeg
# Optional
brew install mecab mecab-ipadic yt-dlp fzf rofi chafa ffmpegthumbnailer
# Optional: subtitle sync (install at least one for subtitle syncing to work)
brew install alass
pip install ffsubsync
```

Grant Accessibility permission to SubMiner in **System Settings > Privacy & Security > Accessibility**.

</details>

<details>
<summary><b>Windows</b></summary>

Install [`mpv`](https://mpv.io/installation/) and [`ffmpeg`](https://ffmpeg.org/download.html) and ensure both are on your `PATH`.

Optionally install [MeCab for Windows](https://taku910.github.io/mecab/#download) with the UTF-8 dictionary for additional metadata enrichment.

</details>

---

## Quick Start

### 1. Install

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
mkdir -p ~/.local/bin
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/SubMiner.AppImage -O ~/.local/bin/SubMiner.AppImage \
	&& chmod +x ~/.local/bin/SubMiner.AppImage
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer -O ~/.local/bin/subminer \
	&& chmod +x ~/.local/bin/subminer
```

> [!NOTE]
> The `subminer` wrapper uses a [Bun](https://bun.sh) shebang. Make sure `bun` is on your `PATH`.

</details>

<details>
<summary><b>macOS</b></summary>

Download the latest DMG or ZIP from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest) and drag `SubMiner.app` into `/Applications`.

Also download the `subminer` launcher (recommended):

```bash
sudo curl -fSL https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer -o /usr/local/bin/subminer \
	&& sudo chmod +x /usr/local/bin/subminer
```

> [!NOTE]
> The `subminer` launcher uses a [Bun](https://bun.sh) shebang. Make sure `bun` is on your `PATH`.

</details>

<details>
<summary><b>Windows</b></summary>

Download the latest installer or portable `.zip` from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest). Make sure `mpv` is on your `PATH`.

> [!WARNING]
> **Windows support is experimental.** Core features such as mining, annotations, and dictionary lookups work, but some functionality may be missing or unstable. Bug reports welcome.

> [!NOTE]
> On Windows the `subminer` launcher requires [`bun`](https://bun.sh) and must be invoked with `bun run subminer` instead of running the script directly. The recommended alternative is the **SubMiner mpv** shortcut created during first-run setup — double-click it, drag files onto it, or run `SubMiner.exe --launch-mpv` from a terminal. See the [Windows mpv Shortcut](https://docs.subminer.moe/usage#windows-mpv-shortcut) section for details.

</details>

<details>
<summary><b>From source</b></summary>

See the [build-from-source guide](https://docs.subminer.moe/installation#from-source).

</details>

### 2. Check Dependencies

```bash
subminer doctor                 # verify mpv, ffmpeg, config, and socket
```

> [!NOTE]
> On Windows, use `bun run subminer doctor` or run `SubMiner.exe` directly — first-run setup will guide you through dependency checks.

### 3. First Launch

Run the app. On first launch SubMiner starts in the system tray, creates a default config, and opens a setup popup to finish config, install the mpv plugin, and configure Yomitan dictionaries.

### 4. Mine

```bash
subminer video.mkv          # play video with overlay
subminer --start video.mkv  # explicit overlay start
subminer stats              # open immersion dashboard
subminer stats -b           # stats daemon in background
subminer stats -s           # stop background stats daemon
```

On **Windows**, the `subminer` script must be run with `bun run subminer` (e.g. `bun run subminer video.mkv`). The recommended alternative is the **SubMiner mpv** shortcut (created during setup) or `SubMiner.exe --launch-mpv`.  Drag a video file onto the shortcut to play it, or double-click it to open mpv with SubMiner's defaults.

## Documentation

Full guides on configuration, Anki setup, Jellyfin, immersion tracking, and more: **[docs.subminer.moe](https://docs.subminer.moe)**

---

## Acknowledgments

SubMiner builds on the work of these open-source projects:

| Project                                                                                     | Role                                                                    |
| ------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------- |
| [Anacreon-Script](https://github.com/friedrich-de/Anacreon-Script)                          | Inspiration for the mining workflow                                     |
| [asbplayer](https://github.com/killergerbah/asbplayer)                                      | Inspiration for subtitle sidebar and logic for YouTube subtitle parsing |
| [Bee's Character Dictionary](https://github.com/bee-san/Japanese_Character_Name_Dictionary) | Character name recognition in subtitles                                 |
| [GameSentenceMiner](https://github.com/bpwhelan/GameSentenceMiner)                          | Inspiration for Electron overlay with Yomitan integration               |
| [jellyfin-mpv-shim](https://github.com/jellyfin/jellyfin-mpv-shim)                          | Jellyfin integration                                                    |
| [Jimaku.cc](https://jimaku.cc)                                                              | Japanese subtitle search and downloads                                  |
| [Renji's Texthooker Page](https://github.com/Renji-XD/texthooker-ui)                        | Base for the WebSocket texthooker integration                           |
| [Yomitan](https://github.com/yomidevs/yomitan)                                              | Dictionary engine powering all lookups and the morphological parser     |
| [yomitan-jlpt-vocab](https://github.com/stephenmk/yomitan-jlpt-vocab)                       | JLPT level tags for vocabulary                                          |

## License

[GNU General Public License v3.0](LICENSE)
