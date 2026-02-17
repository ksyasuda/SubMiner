<div align="center">
  <img src="assets/SubMiner.png" width="169" alt="SubMiner logo">
  <h1>SubMiner</h1>
</div>

An all-in-one immersion mining overlay for MPV with AnkiConnect and dictionary (Yomitan) integration.

## What This Project Is For

SubMiner is for Japanese learners who watch subtitled content in mpv and want a low-friction mining loop:

- stay inside the video while doing lookups
- mine to Anki quickly without manual copy/paste workflows
- preserve card context (sentence, audio, screenshot, translation, metadata)
- reduce tool switching between player, dictionary, and card workflow

## Project Goals

1. Keep immersion continuous by making lookup and mining happen over mpv subtitles.
2. Preserve card quality with context-rich media and subtitle timing.
3. Support real daily workflows (subtitle management, sync, known-word awareness, keyboard-first controls).
4. Stay configurable with sensible defaults and advanced customization.
5. Evolve quickly and safely with a TypeScript codebase and automated tests that make refactors easier to ship.

## Who It's For

- learners using mpv as their primary immersion player
- users already working with Yomitan and AnkiConnect
- miners who care about long-term card quality, not just quick word capture

SubMiner is likely overkill if you only want lightweight dictionary lookup without card enrichment or integrated workflow tools.

## Features

- Real-time subtitle display from MPV via IPC socket
- Yomitan integration for fast, on-screen lookups
- Japanese text tokenization using MeCab with smart word boundary detection
- Integrated texthooker-ui server for use with Yomitan
- Integrated WebSocket server (if [mpv_websocket](https://github.com/kuroahna/mpv_websocket) is not found) to send lines to the texthooker
- AnkiConnect integration for automatic card creation with media (audio/image)
- Invisible subtitle position edit mode (`Ctrl/Cmd+Shift+P`, arrow keys to adjust, `Enter`/`Ctrl+S` save, `Esc` cancel)

## Demo

[![Demo screenshot](./assets/demo-poster.jpg)](https://github.com/user-attachments/assets/9235a554-ea51-4284-b14b-7bbf3defaf58)

## Requirements

- `mpv` with IPC socket support
- `mecab` and `mecab-ipadic` (recommended for Japanese tokenization)
- Linux: Hyprland (`hyprctl`) or X11 (`xdotool` + `xwininfo`)
- macOS: Accessibility permission for window tracking

Optional but recommended: `yt-dlp`, `fzf`, `rofi`, `chafa`, `ffmpegthumbnailer`.

## Install

### Linux (AppImage)

Download the latest release from [GitHub Releases](https://github.com/ksyasuda/SubMiner/releases/latest):

```bash
wget https://github.com/ksyasuda/SubMiner/releases/download/v0.1.0/SubMiner-0.1.0.AppImage -O ~/.local/bin/SubMiner.AppImage
chmod +x ~/.local/bin/SubMiner.AppImage
wget https://github.com/ksyasuda/SubMiner/releases/download/v0.1.0/subminer -O ~/.local/bin/subminer
chmod +x ~/.local/bin/subminer
```

The `subminer` wrapper uses a [Bun](https://bun.sh) shebang, so `bun` must be on `PATH`.

### From Source

```bash
git clone --recurse-submodules https://github.com/ksyasuda/SubMiner.git
cd SubMiner
make build
make install
```

If you already cloned without submodules:

```bash
cd SubMiner
git submodule update --init --recursive
```

For macOS builds, signing, and platform-specific details, see [docs/installation.md](docs/installation.md).

## Quick Start

1. Copy and customize [`config.example.jsonc`](config.example.jsonc) to `$XDG_CONFIG_HOME/SubMiner/config.jsonc` (or `~/.config/SubMiner/config.jsonc`).
2. Start mpv with IPC enabled:

```bash
mpv --input-ipc-server=/tmp/subminer-socket video.mkv
```

3. Launch SubMiner:

```bash
subminer video.mkv
# or
subminer https://youtu.be/...
```

## Common Commands

```bash
subminer                          # pick video from current dir (fzf)
subminer -R                       # use rofi picker
subminer -d ~/Videos              # set source directory
subminer -r -d ~/Anime            # recursive search
subminer video.mkv                # launch with default mpv profile (subminer)
subminer -p gpu-hq video.mkv      # override mpv profile
subminer -T video.mkv             # disable texthooker
```

## MPV Plugin (Optional)

```bash
cp plugin/subminer.lua ~/.config/mpv/scripts/
cp plugin/subminer.conf ~/.config/mpv/script-opts/
# or: make install-plugin
```

Requires mpv IPC: `--input-ipc-server=/tmp/subminer-socket`

Default chord prefix: `y` (`y-y` menu, `y-s` start, `y-S` stop, `y-t` toggle visible layer).
Overlay Jimaku shortcut default: `Ctrl+Shift+J` (`shortcuts.openJimaku`).

## Documentation

Detailed guides live in [`docs/`](docs/README.md):

- [Installation](docs/installation.md) — Platform requirements, AppImage/macOS/source installs, mpv plugin
- [Usage](docs/usage.md) — Script vs plugin workflow, keybindings, YouTube playback
- [Mining Workflow](docs/mining-workflow.md) — End-to-end mining guide, overlay layers, card creation
- [Configuration](docs/configuration.md) — Full config reference and option details
- [Anki Integration](docs/anki-integration.md) — AnkiConnect setup, field mapping, media generation
- [MPV Plugin](docs/mpv-plugin.md) — Chord keybindings, subminer.conf options, script messages
- [Troubleshooting](docs/troubleshooting.md) — Common issues and solutions
- [Development](docs/development.md) — Building, testing, contributing
- [Architecture](docs/architecture.md) — Service-oriented design, composition model, and modular renderer layout (`src/renderer/{modals,handlers,utils,...}`)

### Third-Party Components

This project includes the following third-party components:

- **[Yomitan](https://github.com/yomidevs/yomitan)** - Pop-up dictionary
- **[texthooker-ui](https://github.com/ksyasuda/texthooker-ui/tree/subminer)** - Texthooker Page
- **[yomitan-jlpt-vocab](https://github.com/stephenmk/yomitan-jlpt-vocab)** - JLPT Yomitan Dictionary
- **[Jiten Frequency Dictionary](https://jiten.moe/)** - Frequency Dictionary

### Acknowledgments

- **[GameSentenceMiner](https://github.com/bpwhelan/GameSentenceMiner)** — Inspiration for the subtitle overlay and Yomitan integration
- **[Jimaku.cc](https://jimaku.cc)** — Japanese subtitle provider

This project cherry-picks features from the following MPV scripts, ported to TypeScript:

- **[mpvacious](https://github.com/Ajatt-Tools/mpvacious)** — Sentence mining, screenshotting, and card updating logic
- **[Anacreon-Script (animecards)](https://github.com/friedrich-de/Anacreon-Script)** — Copy/paste to card update flow
- **[autosubsync-mpv](https://github.com/joaquintorres/autosubsync-mpv)** — Subtitle synchronization


## License

GNU General Public License v3.0. See [LICENSE](LICENSE).
