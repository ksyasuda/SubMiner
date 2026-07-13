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
- `assets/themes/subminer.rasi` next to the launcher script (final fallback)

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

`subminer sync <host>` merges immersion stats and watch history between two machines over SSH, so both end up with the union of sessions, lifetime totals, vocabulary counts, daily/monthly charts, and `--history` entries. `<host>` is anything `ssh` accepts (`user@hostname` or an ssh config alias); SubMiner must be installed on both machines at the same version. The sync engine runs only inside the app (`SubMiner --sync-cli sync ...`): the sync window spawns it that way, `subminer sync` is a thin proxy that forwards to the installed app, and the remote side is found automatically whether it has the launcher or just the app — so the command-line launcher is optional everywhere.

```bash
subminer sync macbook                  # two-way sync with the host "macbook"
subminer sync macbook --push           # merge local data into macbook only
subminer sync macbook --pull           # merge macbook data into local only
subminer sync user@192.168.1.20       # explicit user@host
subminer sync macbook --remote-cmd ~/bin/subminer  # custom remote SubMiner/launcher path
subminer sync macbook --check          # test SSH + remote SubMiner without syncing
subminer sync --ui                     # open the sync window (also in the tray menu)
```

How it works: each side takes a consistent snapshot of its database (`VACUUM INTO`), the snapshots are exchanged over `scp`, and each machine merges the other's snapshot into its own database. The merge is an insert-only union keyed on stable identifiers (session UUIDs, video keys, series title keys, word/kanji identity), so it is safe to re-run at any time — syncing twice changes nothing, and nothing is ever overwritten or summed twice. Lifetime totals and rollup charts are updated incrementally, so history older than the session retention window is preserved on both sides.

For a one-way transfer, `--push` snapshots the local database and merges it into the host without changing the local database. `--pull` snapshots the host and merges it into the local database without changing the host. These modes add missing data; they do not delete destination-only data or make the destination an exact mirror.

Close SubMiner (and stop the background stats daemon, `subminer stats -s`) on both machines before syncing; the command refuses to run while a SubMiner process may be writing the database (`--force` overrides). The mpv safety check requires a live socket connection, so a stale socket file left after mpv exits does not block sync. Both machines must be on the same SubMiner version — the sync aborts on a stats schema mismatch.

On the remote, sync looks for the `subminer` launcher first (PATH and `~/.local/bin`), then the app binary in `--sync-cli` mode (`SubMiner` on PATH, then the standard macOS `/Applications` and `~/Applications` installs), checking standard SubMiner and Bun locations (`~/.local/bin`, `~/.bun/bin`, Homebrew, `/usr/local/bin`, `/usr/bin`, and `/bin`) even when the non-interactive SSH shell omits them from `PATH`. An AppImage in a custom location can be addressed with `--remote-cmd /path/to/SubMiner.AppImage` (or symlink it as `SubMiner` somewhere on the remote PATH).

Windows remotes are supported: enable Windows' built-in **OpenSSH Server** and sync detects the remote shell (cmd or PowerShell) automatically, finding SubMiner in its default install location (`%LOCALAPPDATA%\Programs\SubMiner`), the launcher shim (`%LOCALAPPDATA%\SubMiner\bin`), or on PATH. Temp files on the remote are created and removed by SubMiner itself (`sync --make-temp` / `--remove-temp`), so no POSIX tools are required on the remote side.

Two lower-level modes are used internally over SSH and also work standalone for manual transfers (e.g. via a USB drive):

```bash
subminer sync --snapshot /tmp/stats.sqlite   # write a consistent snapshot of the local database
subminer sync --merge /tmp/stats.sqlite      # merge a snapshot file into the local database
```

Unfinished sessions (a crash mid-playback) are skipped until the app finalizes them; they sync on the next run. Word/kanji "known" state from Anki is not part of the database and does not sync — each machine derives it from its own Anki collection.

`subminer sync <host> --check` verifies a host without touching any data: it probes the SSH connection, locates SubMiner on the remote (launcher or app binary), and reports its version. `--json` switches any sync mode to machine-readable NDJSON progress output (this is what the sync window consumes).

`sync --make-temp` creates a restricted temporary directory and prints its path; `sync --remove-temp <dir>` removes one created by that command. They are internal SSH transfer helpers, exposed for compatibility but normally invoked only by sync itself. `SubMiner --sync-cli sync ...` is the packaged app's headless compatibility entrypoint; use `SubMiner --sync-cli --help` for its sync-specific help. The `subminer sync` launcher command selects this entrypoint automatically and runs AppImages in Node-only mode, so remote sync does not require a graphical session.

### Sync window

`subminer sync --ui` opens a dedicated window for the same engine in a detached app process, returning the shell immediately. Closing that standalone-launched window exits its app instance. Opening **Sync Stats & History** from the tray keeps the resident app running when the window closes:

- **Devices** — saved hosts with a per-host direction (two-way / push / pull), an auto-sync toggle, last-sync status, and one-click **Sync now** / **Test** / **Remove**. Hosts synced from the command line appear here automatically.
- **Add a device** — test SSH + remote SubMiner availability before saving, with a setup checklist for first-time SSH configuration.
- **Activity** — live stage-by-stage progress, remote output, and a merge summary (sessions, words, kanji, rollups) when a run finishes. Runs can be cancelled, and guard failures offer a one-click `--force` retry.
- **Snapshots** — create manual database snapshots (stored in `/tmp/subminer-db-snapshots/` by default), merge a snapshot file into the local database, or reveal/delete existing snapshots.

Hosts with **Auto-sync** enabled are synced in the background on a configurable interval (default every 60 minutes) whenever no mpv session or stats server is using the database; results surface as overlay notifications. Host bookkeeping lives in `<config dir>/sync-hosts.json`.

## Common Commands

```bash
subminer video.mkv                      # play a specific file (managed launches auto-start the visible overlay by default)
subminer https://youtu.be/...           # YouTube playback (requires yt-dlp)
subminer --backend x11 video.mkv        # Force x11 backend for a specific file
subminer -u                             # check for SubMiner updates
subminer logs -e                        # export sanitized log ZIP
subminer stats                          # open immersion dashboard
subminer stats -b                       # start background stats daemon
```

## Subcommands

| Subcommand                                 | Purpose                                                                                           |
| ------------------------------------------ | ------------------------------------------------------------------------------------------------- |
| `subminer jellyfin` / `jf`                 | Jellyfin workflows (`-d` discovery, `-p` play, `-l` login, `--logout`, `--setup`)                 |
| `subminer stats`                           | Start the stats server (opens the dashboard when `stats.autoOpenBrowser` is on)                   |
| `subminer stats -b` / `-s`                 | Start/reuse or stop the background stats daemon                                                   |
| `subminer stats cleanup`                   | Backfill vocabulary metadata and prune stale rows (`-v` vocab, `-l` lifetime summaries)           |
| `subminer stats rebuild` / `backfill`      | Rebuild or backfill rollup data                                                                   |
| `subminer doctor`                          | Dependency + config + socket diagnostics (`--refresh-known-words` refreshes the known-word cache) |
| `subminer settings`                        | Open the SubMiner settings window                                                                 |
| `subminer logs -e`                         | Export a sanitized local-date log ZIP and print its path                                          |
| `subminer config path`                     | Print active config file path                                                                     |
| `subminer config show`                     | Print active config contents                                                                      |
| `subminer mpv status`                      | Check mpv socket readiness                                                                        |
| `subminer mpv socket`                      | Print active socket path                                                                          |
| `subminer mpv idle`                        | Launch detached idle mpv instance                                                                 |
| `subminer sync <host>`                     | Two-way stats/history sync with another machine over SSH                                          |
| `subminer sync <host> --push`              | Merge local stats/history into another machine only                                               |
| `subminer sync <host> --pull`              | Merge another machine's stats/history into the local database only                                |
| `subminer sync <host> --check`             | Test SSH connection and remote launcher availability                                              |
| `subminer sync --ui`                       | Open the sync window (saved devices, auto-sync, snapshots)                                        |
| `subminer dictionary <path>` / `dict`      | Generate character dictionary ZIP from file/dir target                                            |
| `subminer dictionary --candidates <path>`  | List AniList candidate matches for character dictionary correction                                |
| `subminer dictionary --select <id> <path>` | Pin an AniList media ID for that target series                                                    |
| `subminer texthooker`                      | Launch texthooker-only mode                                                                       |
| `subminer texthooker -o`                   | Launch texthooker and open it in the default browser                                              |
| `subminer app` / `bin`                     | Pass arguments directly to SubMiner binary (e.g. `subminer app --setup`)                          |

Use `subminer <subcommand> -h` for command-specific help.

## Options

| Flag                  | Description                                                                 |
| --------------------- | --------------------------------------------------------------------------- |
| `-d, --directory`     | Video search directory (default: cwd)                                       |
| `-r, --recursive`     | Search directories recursively                                              |
| `-R, --rofi`          | Use rofi instead of fzf                                                     |
| `-H, --history`       | Browse local watch history (see [Watch History](#watch-history))            |
| `-v, --version`       | Print the launcher's own version (can differ from the installed app binary) |
| `-u, --update`        | Check for SubMiner updates and update the app/launcher when possible        |
| `--start`             | Explicitly start overlay after mpv launches                                 |
| `-S, --start-overlay` | Force the visible overlay on start                                          |
| `-T, --no-texthooker` | Disable texthooker server                                                   |
| `-p, --profile`       | mpv profile name (no default; omitted unless set)                           |
| `-a, --args`          | Pass additional mpv arguments as a quoted string                            |
| `-b, --backend`       | Force window backend (`hyprland`, `sway`, `x11`, `macos`, `windows`)        |
| `--settings`          | Open the SubMiner settings window                                           |
| `--log-level`         | Logger verbosity (`debug`, `info`, `warn`, `error`)                         |

App-binary flags such as `--setup`, `--dev`, and `--debug` are not launcher flags - pass them through with `subminer app`, for example `subminer app --setup`.

On Linux, `subminer -u` updates from the launcher process itself. It can check and replace the AppImage, launcher, runtime plugin copy, and rofi theme even when SubMiner is already running in the tray.

Managed launches inject `auto_start=yes`, `auto_start_visible_overlay=yes`, and `auto_start_pause_until_ready=yes` as plugin script-opts from SubMiner's config defaults (`mpv.autoStartSubMiner`, `auto_start_overlay`), so explicit start flags are usually unnecessary. The plugin's own built-in defaults are off - mpv launched outside SubMiner does not auto-start the overlay.

## Logging

- Default log level is `warn` (launcher and app; configurable via `logging.level`)
- `--dev` / `--debug` are app-binary flags that control app dev-mode, not logging verbosity - use `--log-level` for that
