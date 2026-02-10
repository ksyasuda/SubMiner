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
- Integrated websocket server (if [mpv_websocket](https://github.com/kuroahna/mpv_websocket) is not found) to send lines to the texthooker
- AnkiConnect integration for automatic card creation with media (audio/image)

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

```bash
wget https://github.com/sudacode/subminer/releases/download/v1.0.0/subminer-1.0.0.AppImage -O ~/.local/bin/subminer.AppImage
chmod +x ~/.local/bin/subminer.AppImage
wget https://github.com/sudacode/subminer/releases/download/v1.0.0/subminer -O ~/.local/bin/subminer
chmod +x ~/.local/bin/subminer
```

`subminer` uses a [Bun](https://bun.com) shebang, so `bun` must be on `PATH`.

### From source

```bash
git clone https://github.com/sudacode/subminer.git
cd subminer
make build
make install
```

For macOS app bundle / signing / permissions details, use `docs/installation.md`.

## Quick Start

1. Copy and customize [`config.example.jsonc`](config.example.jsonc) to `~/.config/SubMiner/config.jsonc`.
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
```

Requires mpv IPC: `--input-ipc-server=/tmp/subminer-socket`

Default chord prefix: `y` (`y-y` menu, `y-s` start, `y-S` stop, `y-t` toggle visible layer).

## Documentation

Detailed guides live in [`docs/`](docs/README.md):

- [Installation](docs/installation.md)
- [Usage](docs/usage.md)
- [Configuration](docs/configuration.md)
- [Development](docs/development.md)

### Third-Party Components

This project includes the following third-party components:

- **[Yomitan](https://github.com/yomidevs/yomitan)** - GPL-3.0
- **[texthooker-ui](https://github.com/Renji-XD/texthooker-ui)** - MIT

### Acknowledgments

This project was inspired by **[GameSentenceMiner](https://github.com/bpwhelan/GameSentenceMiner)**'s subtitle overlay and Yomitan integration

## License

GNU General Public License v3.0. See [LICENSE](LICENSE).
