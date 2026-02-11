<div align="center">
  <img src="assets/SubMiner.png" width="169" alt="SubMiner logo">
  <h1>SubMiner</h1>
</div>

An all-in-one sentence mining overlay for MPV with AnkiConnect and dictionary (Yomitan) integration.

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
git clone https://github.com/ksyasuda/SubMiner.git
cd SubMiner
make build
make install
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
Overlay Jimaku shortcut default: `Ctrl+Alt+J` (`shortcuts.openJimaku`).

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
- [Architecture](docs/architecture.md) — Service-oriented design, composition model

### Third-Party Components

This project includes the following third-party components:

- **[Yomitan](https://github.com/yomidevs/yomitan)** — GPL-3.0
- **[texthooker-ui](https://github.com/Renji-XD/texthooker-ui)** — MIT

### Acknowledgments

This project was inspired by **[GameSentenceMiner](https://github.com/bpwhelan/GameSentenceMiner)**'s subtitle overlay and Yomitan integration

## License

GNU General Public License v3.0. See [LICENSE](LICENSE).
