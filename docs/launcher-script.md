# Launcher Script

The `subminer` wrapper script is an all-in-one launcher that handles video selection, mpv startup, and overlay management. It's a Bun script distributed alongside the AppImage.

## Video Picker

When you run `subminer` without specifying a file, it opens an interactive video picker. By default it uses **fzf** in the terminal; pass `-R` to use **rofi** instead.

### fzf (default)

```bash
subminer                        # pick from current directory
subminer -d ~/Videos            # pick from a specific directory
subminer -r -d ~/Anime          # recursive search
```

fzf shows video files in a fuzzy-searchable list. If `chafa` is installed, you get thumbnail previews in the right pane. Thumbnails are sourced from the freedesktop thumbnail cache first, then generated on the fly with `ffmpegthumbnailer` or `ffmpeg` as fallback.

| Optional tool         | Purpose                          |
| --------------------- | -------------------------------- |
| `chafa`               | Render thumbnails in the terminal |
| `ffmpegthumbnailer`   | Generate thumbnails on the fly   |

### rofi

```bash
subminer -R                     # rofi picker, current directory
subminer -R -d ~/Videos         # rofi picker, specific directory
subminer -R -r -d ~/Anime       # rofi picker, recursive
```

rofi shows a GUI menu with icon thumbnails when available. SubMiner ships a custom rofi theme that can be installed from the release assets:

```bash
mkdir -p ~/.local/share/SubMiner/themes
cp /tmp/assets/themes/subminer.rasi ~/.local/share/SubMiner/themes/subminer.rasi
```

The theme is auto-detected from these paths (first match wins):

- `$SUBMINER_ROFI_THEME` environment variable (absolute path)
- `$XDG_DATA_HOME/SubMiner/themes/subminer.rasi` (default: `~/.local/share/SubMiner/themes/subminer.rasi`)
- `/usr/local/share/SubMiner/themes/subminer.rasi`
- `/usr/share/SubMiner/themes/subminer.rasi`
- macOS: `~/Library/Application Support/SubMiner/themes/subminer.rasi`

Override with the `SUBMINER_ROFI_THEME` environment variable:

```bash
SUBMINER_ROFI_THEME=/path/to/custom-theme.rasi subminer -R
```

## Common Commands

```bash
subminer video.mkv              # play a specific file
subminer --start video.mkv      # play + explicitly start overlay
subminer https://youtu.be/...   # YouTube playback (requires yt-dlp)
subminer ytsearch:"jp news"     # YouTube search
```

## Subcommands

| Subcommand                | Purpose                                        |
| ------------------------- | ---------------------------------------------- |
| `subminer jellyfin` / `jf` | Jellyfin workflows (`-d` discovery, `-p` play, `-l` login) |
| `subminer yt` / `youtube` | YouTube shorthand (`-o`, `-m`)                  |
| `subminer doctor`         | Dependency + config + socket diagnostics        |
| `subminer config path`    | Print active config file path                   |
| `subminer config show`    | Print active config contents                    |
| `subminer mpv status`     | Check mpv socket readiness                      |
| `subminer mpv socket`     | Print active socket path                        |
| `subminer mpv idle`       | Launch detached idle mpv instance               |
| `subminer texthooker`     | Launch texthooker-only mode                     |
| `subminer app`            | Pass arguments directly to SubMiner binary      |

Use `subminer <subcommand> -h` for command-specific help.

## Options

| Flag                 | Description                                  |
| -------------------- | -------------------------------------------- |
| `-d, --directory`    | Video search directory (default: cwd)        |
| `-r, --recursive`    | Search directories recursively               |
| `-R, --rofi`         | Use rofi instead of fzf                      |
| `-S, --start`        | Start overlay after mpv launches             |
| `-T, --no-texthooker`| Disable texthooker server                    |
| `-p, --profile`      | mpv profile name (default: `subminer`)       |
| `-b, --backend`      | Force window backend (`hyprland`, `sway`, `x11`) |
| `--log-level`        | Logger verbosity (`debug`, `info`, `warn`, `error`) |
| `--dev`, `--debug`   | Enable app dev-mode (not tied to log level)  |

## Logging

- Default log level is `info`
- `--background` mode defaults to `warn` unless `--log-level` is explicitly set
- `--dev` / `--debug` control app behavior, not logging verbosity — use `--log-level` for that

