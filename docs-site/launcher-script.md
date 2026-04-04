# Launcher Script

The `subminer` launcher is an all-in-one script that handles video selection, mpv startup, and overlay management. It is the recommended way to use SubMiner on Linux and macOS because it guarantees mpv is launched with the correct IPC socket and SubMiner defaults. It's a Bun script distributed as a release asset alongside the AppImage and DMG.

::: tip Windows users
On Windows the `subminer` script cannot run directly via shebang — use `bun run subminer` instead (e.g. `bun run subminer video.mkv`). The recommended alternative is the **SubMiner mpv** shortcut created during first-run setup, or `SubMiner.exe --launch-mpv`. See [Windows mpv Shortcut](/usage#windows-mpv-shortcut) for details.
:::

## Video Picker

When you run `subminer` without specifying a file, it opens an interactive video picker. By default it uses **fzf** in the terminal; pass `-R` to use **rofi** instead.

### fzf (default)

```bash
subminer                        # pick from current directory
subminer -d ~/Videos            # pick from a specific directory
subminer -r -d ~/Anime          # recursive search
```

fzf shows video files in a fuzzy-searchable list. If `chafa` is installed, you get thumbnail previews in the right pane. Thumbnails are sourced from the freedesktop thumbnail cache first, then generated on the fly with `ffmpegthumbnailer` or `ffmpeg` as fallback.

| Optional tool       | Purpose                           |
| ------------------- | --------------------------------- |
| `chafa`             | Render thumbnails in the terminal |
| `ffmpegthumbnailer` | Generate thumbnails on the fly    |

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
subminer video.mkv              # play a specific file (default plugin config auto-starts visible overlay)
subminer --start video.mkv      # optional explicit overlay start when plugin auto_start=no
subminer -S video.mkv           # same as above via --start-overlay
subminer https://youtu.be/...   # YouTube playback (requires yt-dlp)
subminer ytsearch:"jp news"     # YouTube search
subminer stats                  # open immersion dashboard
subminer stats -b               # start background stats daemon
subminer stats -s               # stop background stats daemon
subminer --setup                # Open first-run setup popup
```

## Subcommands

| Subcommand                   | Purpose                                                    |
| ---------------------------- | ---------------------------------------------------------- |
| `subminer jellyfin` / `jf`   | Jellyfin workflows (`-d` discovery, `-p` play, `-l` login) |
| `subminer stats`             | Start stats server and open immersion dashboard in browser |
| `subminer stats -b`          | Start or reuse background stats daemon (non-blocking)      |
| `subminer stats -s`          | Stop the background stats daemon                           |
| `subminer stats cleanup`     | Backfill vocabulary metadata and prune stale rows          |
| `subminer doctor`            | Dependency + config + socket diagnostics                   |
| `subminer config path`       | Print active config file path                              |
| `subminer config show`       | Print active config contents                               |
| `subminer mpv status`        | Check mpv socket readiness                                 |
| `subminer mpv socket`        | Print active socket path                                   |
| `subminer mpv idle`          | Launch detached idle mpv instance                          |
| `subminer dictionary <path>` | Generate character dictionary ZIP from file/dir target     |
| `subminer texthooker`        | Launch texthooker-only mode                                |
| `subminer app`               | Pass arguments directly to SubMiner binary                 |

Use `subminer <subcommand> -h` for command-specific help.

## Options

| Flag                  | Description                                         |
| --------------------- | --------------------------------------------------- |
| `-d, --directory`     | Video search directory (default: cwd)               |
| `-r, --recursive`     | Search directories recursively                      |
| `-R, --rofi`          | Use rofi instead of fzf                             |
| `--setup`               | Open first-run setup popup manually                    |
| `--start`             | Explicitly start overlay after mpv launches         |
| `-S, --start-overlay` | Explicitly start overlay after mpv launches         |
| `-T, --no-texthooker` | Disable texthooker server                           |
| `-p, --profile`       | mpv profile name (default: `subminer`)              |
| `-a, --args`          | Pass additional mpv arguments as a quoted string      |
| `-b, --backend`       | Force window backend (`hyprland`, `sway`, `x11`, `macos`, `windows`)    |
| `--log-level`         | Logger verbosity (`debug`, `info`, `warn`, `error`) |
| `--dev`, `--debug`    | Enable app dev-mode (not tied to log level)         |

With default plugin settings (`auto_start=yes`, `auto_start_visible_overlay=yes`, `auto_start_pause_until_ready=yes`), explicit start flags are usually unnecessary.

## Logging

- Default log level is `info`
- `--background` mode defaults to `warn` unless `--log-level` is explicitly set
- `--dev` / `--debug` control app behavior, not logging verbosity — use `--log-level` for that
