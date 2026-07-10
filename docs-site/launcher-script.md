# Launcher Script

The `subminer` launcher is an all-in-one script that handles video selection, mpv startup, and overlay management. It is the recommended way to use SubMiner on Linux and macOS because it guarantees mpv is launched with the correct IPC socket and SubMiner defaults. It's a Bun script distributed as a release asset alongside the AppImage and DMG.

::: tip Windows users
On Windows, the recommended way to launch playback is the **SubMiner mpv** shortcut created during first-run setup - double-click it, drag a file onto it, or run `SubMiner.exe --launch-mpv` from a terminal. See [Windows mpv Shortcut](/usage#windows-mpv-shortcut) for details.
:::

## Video Picker

When you run `subminer` without specifying a file, it opens an interactive video picker. By default it uses **fzf** in the terminal; pass `-R` to use **rofi** instead.

### fzf (default)

```bash
subminer                               # pick from current directory
subminer -d ~/Videos                   # pick from a specific directory
subminer -r -d ~/Anime                 # recursive search
```

fzf shows video files in a fuzzy-searchable list. If `chafa` is installed, you get thumbnail previews in the right pane. Thumbnails are sourced from the freedesktop thumbnail cache first, then generated on the fly with `ffmpegthumbnailer` or `ffmpeg` as fallback.

| Optional tool       | Purpose                           |
| ------------------- | --------------------------------- |
| `chafa`             | Render thumbnails in the terminal |
| `ffmpegthumbnailer` | Generate thumbnails on the fly    |

### rofi

```bash
subminer -R                            # rofi picker, current directory
subminer -R -d ~/Videos                # rofi picker, specific directory
subminer -R -r -d ~/Anime              # rofi picker, recursive
subminer -R /directory                 # rofi picker, directory shortcut
```

rofi shows a GUI menu with icon thumbnails when available. SubMiner ships the rofi theme plus the Linux launcher-managed runtime plugin copy in the release assets tarball:

```bash
wget https://github.com/ksyasuda/SubMiner/releases/latest/download/subminer-assets.tar.gz -O /tmp/subminer-assets.tar.gz
tar -xzf /tmp/subminer-assets.tar.gz -C /tmp
mkdir -p ~/.local/share/SubMiner/themes
cp /tmp/assets/themes/subminer.rasi ~/.local/share/SubMiner/themes/subminer.rasi
mkdir -p ~/.local/share/SubMiner/plugin
cp -R /tmp/plugin/subminer ~/.local/share/SubMiner/plugin/subminer
```

Once the `SubMiner` data dir exists, `subminer -u` refreshes both assets automatically. Normal Linux launcher playback also checks for the managed runtime plugin copy and rofi theme before mpv launch and installs them from the bundled app automatically if either one is missing.

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

## Watch History

`subminer -H` (or `--history`) browses your local watch history, sourced from the immersion tracker database. It works with both pickers: fzf by default, rofi with `-R -H`.

```bash
subminer -H                            # fzf history browser
subminer -R -H                         # rofi history browser
```

The first menu lists every locally watched series, most recently watched first, using the parsed media title (e.g. the anime title) when available and the directory name otherwise. Selecting a series opens an action menu:

- **Replay last watched** — replays the most recently watched episode
- **Next episode** — plays the episode after the last watched one (continues into the next season directory when the season ends)
- **Browse episodes** — lists the video files in the series directory in episode order, using the same fzf/rofi episode picker as directory browsing; if the series has multiple season directories, a season menu is shown first

Series whose directories are not currently accessible (e.g. an unmounted network share) are hidden from the list. Watch history requires the immersion tracker database (`immersionTracking.dbPath`, default `<config dir>/immersion.sqlite`), which SubMiner populates during playback.

## Sync Between Machines

`subminer sync <host>` merges immersion stats and watch history between two machines over SSH, so both end up with the union of sessions, lifetime totals, vocabulary counts, daily/monthly charts, and `--history` entries. `<host>` is anything `ssh` accepts (`user@hostname` or an ssh config alias); SubMiner must be installed on both machines at the same version.

```bash
subminer sync macbook                  # two-way sync with the host "macbook"
subminer sync user@192.168.1.20       # explicit user@host
subminer sync macbook --remote-cmd ~/bin/subminer  # custom remote launcher path
```

How it works: each side takes a consistent snapshot of its database (`VACUUM INTO`), the snapshots are exchanged over `scp`, and each machine merges the other's snapshot into its own database. The merge is an insert-only union keyed on stable identifiers (session UUIDs, video keys, series title keys, word/kanji identity), so it is safe to re-run at any time — syncing twice changes nothing, and nothing is ever overwritten or summed twice. Lifetime totals and rollup charts are updated incrementally, so history older than the session retention window is preserved on both sides.

Close SubMiner (and stop the background stats daemon, `subminer stats -s`) on both machines before syncing; the command refuses to run while a SubMiner process may be writing the database (`--force` overrides). Both machines must be on the same SubMiner version — the sync aborts on a stats schema mismatch. Remote sync checks standard SubMiner and Bun locations (`~/.local/bin`, `~/.bun/bin`, Homebrew, `/usr/local/bin`, `/usr/bin`, and `/bin`) even when the non-interactive SSH shell omits them from `PATH`.

Two lower-level modes are used internally over SSH and also work standalone for manual transfers (e.g. via a USB drive):

```bash
subminer sync --snapshot /tmp/stats.sqlite   # write a consistent snapshot of the local database
subminer sync --merge /tmp/stats.sqlite      # merge a snapshot file into the local database
```

Unfinished sessions (a crash mid-playback) are skipped until the app finalizes them; they sync on the next run. Word/kanji "known" state from Anki is not part of the database and does not sync — each machine derives it from its own Anki collection.

## Common Commands

```bash
subminer video.mkv                      # play a specific file (default plugin config auto-starts visible overlay)
subminer https://youtu.be/...           # YouTube playback (requires yt-dlp)
subminer --backend x11 video.mkv        # Force x11 backend for a specific file
subminer -u                             # check for SubMiner updates
subminer logs -e                        # export sanitized log ZIP
subminer stats                          # open immersion dashboard
subminer stats -b                       # start background stats daemon
```

## Subcommands

| Subcommand                                 | Purpose                                                            |
| ------------------------------------------ | ------------------------------------------------------------------ |
| `subminer jellyfin` / `jf`                 | Jellyfin workflows (`-d` discovery, `-p` play, `-l` login)         |
| `subminer stats`                           | Start stats server and open immersion dashboard in browser         |
| `subminer stats -b`                        | Start or reuse background stats daemon (non-blocking)              |
| `subminer stats cleanup`                   | Backfill vocabulary metadata and prune stale rows                  |
| `subminer doctor`                          | Dependency + config + socket diagnostics                           |
| `subminer settings`                        | Open the SubMiner settings window                                  |
| `subminer logs -e`                         | Export a sanitized local-date log ZIP and print its path           |
| `subminer config path`                     | Print active config file path                                      |
| `subminer config show`                     | Print active config contents                                       |
| `subminer mpv status`                      | Check mpv socket readiness                                         |
| `subminer mpv socket`                      | Print active socket path                                           |
| `subminer mpv idle`                        | Launch detached idle mpv instance                                  |
| `subminer sync <host>`                     | Two-way stats/history sync with another machine over SSH           |
| `subminer dictionary <path>`               | Generate character dictionary ZIP from file/dir target             |
| `subminer dictionary --candidates <path>`  | List AniList candidate matches for character dictionary correction |
| `subminer dictionary --select <id> <path>` | Pin an AniList media ID for that target series                     |
| `subminer texthooker`                      | Launch texthooker-only mode                                        |
| `subminer texthooker -o`                   | Launch texthooker and open it in the default browser               |
| `subminer app`                             | Pass arguments directly to SubMiner binary                         |

Use `subminer <subcommand> -h` for command-specific help.

## Options

| Flag                  | Description                                                          |
| --------------------- | -------------------------------------------------------------------- |
| `-d, --directory`     | Video search directory (default: cwd)                                |
| `-r, --recursive`     | Search directories recursively                                       |
| `-R, --rofi`          | Use rofi instead of fzf                                              |
| `-H, --history`       | Browse local watch history (see [Watch History](#watch-history))     |
| `--setup`             | Open first-run setup popup manually                                  |
| `-v, --version`       | Print installed SubMiner version                                     |
| `-u, --update`        | Check for SubMiner updates and update the app/launcher when possible |
| `--start`             | Explicitly start overlay after mpv launches                          |
| `-S, --start-overlay` | Explicitly start overlay after mpv launches                          |
| `-T, --no-texthooker` | Disable texthooker server                                            |
| `-p, --profile`       | mpv profile name (no default; omitted unless set)                    |
| `-a, --args`          | Pass additional mpv arguments as a quoted string                     |
| `-b, --backend`       | Force window backend (`hyprland`, `sway`, `x11`, `macos`, `windows`) |
| `--log-level`         | Logger verbosity (`debug`, `info`, `warn`, `error`)                  |
| `--dev`, `--debug`    | Enable app dev-mode (not tied to log level)                          |

On Linux, `subminer -u` updates from the launcher process itself. It can check and replace the AppImage, launcher, runtime plugin copy, and rofi theme even when SubMiner is already running in the tray.

With default plugin settings (`auto_start=yes`, `auto_start_visible_overlay=yes`, `auto_start_pause_until_ready=yes`), explicit start flags are usually unnecessary.

## Logging

- Default log level is `info`
- `--background` mode defaults to `warn` unless `--log-level` is explicitly set
- `--dev` / `--debug` control app behavior, not logging verbosity - use `--log-level` for that
