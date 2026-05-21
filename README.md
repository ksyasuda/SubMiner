<div align="center">

<img src="assets/SubMiner.png" width="160" alt="SubMiner logo">

# SubMiner

Integrates Yomitan and mpv - on-screen lookups, mine to Anki, and track immersion without leaving the player

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

Hover over any word and trigger a lookup to get the full Yomitan popup - definitions, pitch accent, and frequency data - without ever leaving mpv.

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

Real-time subtitle annotations with frequency highlighting, JLPT tags, N+1 targeting, and a character name dictionary. Grammar-only tokens and particles render as plain text so you focus on what matters.

<div align="center">
  <img src="docs-site/public/screenshots/annotations.png" width="800" alt="Annotated subtitles with frequency coloring, JLPT underlines, and N+1 targets">
</div>

<br>

### Immersion Dashboard

Local stats dashboard tracking watch time, vocabulary growth, mining throughput, session history, and trends. All stored locally, no third-party tracking.

<div align="center">
  <img src="docs-site/public/screenshots/stats-overview.png" width="800" alt="Stats dashboard showing watch time, cards mined, streaks, and tracking data">
</div>

<br>

### Playlist Browser

Browse sibling episode files and the active mpv queue in one overlay modal. Open it with `Ctrl+Alt+P` to append episodes from the current directory, jump to queued items, remove entries, or reorder the playlist without leaving playback.

<div align="center">
  <img src="docs-site/public/screenshots/playlist-browser.png" width="800" alt="Stats dashboard showing watch time, cards mined, streaks, and tracking data">
</div>

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
    <td>Browse, launch, and cast media from your Jellyfin server with setup and discovery controls in the app tray</td>
  </tr>
  <tr>
    <td><b>Jimaku</b></td>
    <td>Search and download Japanese subtitles</td>
  </tr>
  <tr>
    <td><b>alass / ffsubsync</b></td>
    <td>Manual subtitle retiming — requires <code>alass</code> or <code>ffsubsync</code> on your <code>PATH</code> (optional; subtitle syncing is disabled without them)</td>
  </tr>
  <tr>
    <td><b>WebSocket</b></td>
    <td>Plain subtitle feed plus a dedicated annotated feed for texthooker pages and custom tools</td>
  </tr>
</table>

<div align="center">
  <img src="docs-site/public/screenshots/texthooker.png" width="800" alt="Texthooker page receiving annotated subtitle lines via WebSocket">
</div>

<br>

---

## Requirements

Only **mpv** and Anki+AnkiConnect are required. Everything else is optional but enhances the experience.

| Dependency           | Status      | What it does                             |
| -------------------- | ----------- | ---------------------------------------- |
| mpv                  | Required    | The video player SubMiner overlays on    |
| Anki + AnkiConnect   | Required    | Card creation from the Yomitan popup     |
| ffmpeg               | Recommended | Audio clips & screenshots for Anki cards |
| MeCab + mecab-ipadic | Recommended | More precise annotations and filtering   |
| yt-dlp               | Optional    | YouTube playback                         |
| fzf / rofi           | Optional    | Video picker in the launcher             |
| alass / ffsubsync    | Optional    | Subtitle sync                            |

<details>
<summary><b>Platform-specific install commands</b></summary>

**Arch Linux:**

```bash
sudo pacman -S --needed mpv ffmpeg mecab mecab-ipadic
```

**macOS:**

```bash
brew install mpv ffmpeg mecab mecab-ipadic
```

**Windows:** Install [mpv](https://mpv.io/installation/) and [ffmpeg](https://ffmpeg.org/download.html) and ensure both are on `PATH`.

See the [full requirements list](https://docs.subminer.moe/installation#1-install-requirements) for optional dependencies.

</details>

---

## Quick Start

### 1. Install SubMiner

<details>
<summary><b>Arch Linux (AUR)</b></summary>

```bash
paru -S subminer-bin
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

</details>

<details>
<summary><b>macOS (DMG)</b></summary>

Download the latest DMG from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest) and drag `SubMiner.app` into `/Applications`.

</details>

<details>
<summary><b>Windows</b></summary>

Download and run the latest installer (`.exe`) from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest).

</details>

<details>
<summary><b>From source</b></summary>

See the [build-from-source guide](https://docs.subminer.moe/installation#from-source).

</details>

### 2. Launch & Set Up

Run SubMiner and the first-run setup wizard will guide you through importing Yomitan dictionaries and optionally installing the `subminer` command-line launcher.

```bash
# Linux
subminer app --setup

# macOS — open SubMiner.app, or:
subminer app --setup
```

On **Windows**, just run `SubMiner.exe` and the setup will open automatically on first launch.

### 3. Mine

```bash
subminer video.mkv          # launch mpv with SubMiner
subminer /path/to/dir       # pick a file with fzf
subminer -R /path/to/dir    # pick a file with rofi (Linux only)
```

On **Windows**, use the **SubMiner mpv** shortcut created during setup. Double-click it or drag a video file onto it.

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
