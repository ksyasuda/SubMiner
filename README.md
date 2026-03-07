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

- **Hover to look up** — Yomitan dictionary popups directly on subtitles
- **Keyboard-driven lookup mode** — Navigate token-by-token, keep lookup open across tokens, and control popup scrolling/audio/mining without leaving the overlay
- **One-key mining** — Creates Anki cards with sentence, audio, screenshot, and translation
- **Instant auto-enrichment** — Optional local AnkiConnect proxy enriches new Yomitan cards immediately
- **Reading annotations** — Combines N+1 targeting, frequency-dictionary highlighting, and JLPT underlining while you read
- **Hover-aware playback** — By default, hovering subtitle text pauses mpv and resumes on mouse leave (`subtitleStyle.autoPauseVideoOnHover`)
- **Subtitle tools** — Download from Jimaku, sync with alass/ffsubsync
- **Immersion tracking** — SQLite-powered stats on your watch time and mining activity
- **Custom texthooker page** — Built-in custom texthooker page and websocket, no extra setup
- **Jellyfin integration** — Remote playback setup, cast device mode, and direct playback launch
- **AniList progress** — Track episode completion and push watching progress automatically

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

**From source** or **macOS** — see the [installation guide](https://docs.subminer.moe/installation#from-source).

### 2. Install the mpv plugin and configuration file

```bash
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer-assets.tar.gz -O /tmp/subminer-assets.tar.gz
tar -xzf /tmp/subminer-assets.tar.gz -C /tmp
mkdir -p ~/.config/mpv/scripts/subminer
mkdir -p ~/.config/mpv/script-opts
cp -R /tmp/plugin/subminer/. ~/.config/mpv/scripts/subminer/
cp /tmp/plugin/subminer.conf ~/.config/mpv/script-opts/
mkdir -p ~/.config/SubMiner && cp /tmp/config.example.jsonc ~/.config/SubMiner/config.jsonc
```

### 3. Set up Yomitan Dictionaries

```bash
subminer app --yomitan
```

### 4. Mine

```bash
subminer app --start --background
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

For full guides on configuration, Anki, Jellyfin, and more, see [docs.subminer.moe](https://docs.subminer.moe).

## Testing

- Run `bun run test` or `bun run test:fast` for the default fast lane: config/core coverage plus representative entry/runtime, Anki integration, and main runtime checks.
- Run `bun run test:full` for the maintained test surface: Bun-compatible `src/**` coverage, Bun-compatible launcher unit coverage, and a Node compatibility lane for suites that depend on Electron named exports or `node:sqlite` behavior.
- Run `bun run test:node:compat` directly when you only need the Node-backed compatibility slice: `ipc`, `anki-jimaku-ipc`, `overlay-manager`, `config-validation`, `startup-config`, and runtime registry coverage.
- Run `bun run test:env` for environment-specific verification: launcher smoke/plugin checks plus the SQLite-backed immersion tracker lane.
- Run `bun run test:immersion:sqlite` when you specifically need real SQLite persistence coverage under Node with `--experimental-sqlite`.
- Run `bun run test:subtitle` for the maintained `alass`/`ffsubsync` subtitle surface.

The Bun-managed discovery lanes intentionally exclude a small set of suites that are currently Node-only because of Bun runtime/tooling gaps rather than product behavior: Electron named-export tests in `src/core/services/ipc.test.ts`, `src/core/services/anki-jimaku-ipc.test.ts`, and `src/core/services/overlay-manager.test.ts`, plus runtime/config tests in `src/main/config-validation.test.ts`, `src/main/runtime/startup-config.test.ts`, and `src/main/runtime/registry.test.ts`. `bun run test:node:compat` keeps those suites in the standard workflow instead of leaving them untracked.

## Acknowledgments

Built on the shoulders of [GameSentenceMiner](https://github.com/bpwhelan/GameSentenceMiner), [Renji's Texthooker Page](https://github.com/Renji-XD/texthooker-ui), [mpvacious](https://github.com/Ajatt-Tools/mpvacious), [Anacreon-Script](https://github.com/friedrich-de/Anacreon-Script), and [Bee's Character Dictionary](https://github.com/bee-san/Japanese_Character_Name_Dictionary). Subtitles powered by [Jimaku.cc](https://jimaku.cc). Dictionary lookups via [Yomitan](https://github.com/yomidevs/yomitan).

## License

[GNU General Public License v3.0](LICENSE)
