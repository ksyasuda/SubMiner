# Launcher script

`subminer` is the command-line entry point for SubMiner. It starts mpv with the socket and options SubMiner needs, opens file pickers, and runs helper commands. This page is the reference for its subcommands and flags. For everyday use, start with [Usage](/usage).

You do not need Bun or anything else installed to run it. It uses the runtime bundled with the app. On Windows, the **SubMiner mpv** shortcut is the simpler way to play files (see [Windows mpv shortcut](/usage#windows-mpv-shortcut)).

```bash
subminer [options] [file | directory | URL]
subminer <subcommand> [options]
```

Run `subminer -h` or `subminer <subcommand> -h` for built-in help.

## Options

| Flag                  | Description                                                                         |
| --------------------- | ----------------------------------------------------------------------------------- |
| `-d, --directory`     | Directory to browse (default: current directory)                                    |
| `-r, --recursive`     | Search subdirectories                                                               |
| `-R, --rofi`          | Use rofi instead of fzf                                                             |
| `-H, --history`       | Browse [watch history](#watch-history)                                              |
| `-b, --backend`       | Window backend: `auto`, `hyprland`, `sway`, `x11`, `macos`, `windows`               |
| `-p, --profile`       | mpv profile to load                                                                 |
| `-a, --args`          | Extra mpv options as one quoted string, e.g. `--args "--volume=80"`                 |
| `--start`             | Start the overlay after mpv launches. Only needed if `mpv.autoStartSubMiner` is off |
| `-S, --start-overlay` | Show the overlay on start                                                           |
| `-T, --no-texthooker` | Do not start the texthooker server                                                  |
| `--settings`          | Open the settings window                                                            |
| `--log-level`         | `debug`, `info`, `warn`, or `error`                                                 |
| `-u, --update`        | Check for and install updates                                                       |
| `-v, --version`       | Print the launcher's version                                                        |

The target can be a video file, a directory (opens the picker there), a URL, or `ytsearch:"query"` for the first YouTube search result.

App flags such as `--setup` and `--dev` are not launcher flags. Pass them through with `subminer app`, for example `subminer app --setup`.

## Subcommands

| Command                                    | What it does                                                                                                         |
| ------------------------------------------ | -------------------------------------------------------------------------------------------------------------------- |
| `subminer stats`                           | Start the stats dashboard server. Opens your browser if `stats.autoOpenBrowser` is on                                |
| `subminer stats -b` / `-s`                 | Start (or reuse) the stats server in the background / stop it                                                        |
| `subminer stats cleanup`                   | Backfill vocabulary metadata and prune stale rows (same as `-v`)                                                     |
| `subminer stats cleanup -l`                | Rebuild lifetime totals from the retained sessions                                                                   |
| `subminer stats cleanup -d`                | Collapse repeated lines from typeset subtitles. Add `--dry-run` to preview, `--lookback-days <n>` to limit the range |
| `subminer stats rebuild` / `backfill`      | Same as `stats cleanup -l`                                                                                           |
| `subminer sync <host>`                     | Sync stats and watch history with another machine. See [below](#sync-between-machines)                               |
| `subminer doctor`                          | Check the app, mpv, ffmpeg, yt-dlp, pickers, config, and mpv socket                                                  |
| `subminer doctor --refresh-known-words`    | Refresh the known-word cache from Anki                                                                               |
| `subminer settings`                        | Open the settings window                                                                                             |
| `subminer generate-subs [video]`           | Generate [Japanese subtitles](/subtitle-generation) with whisper.cpp                                                 |
| `subminer jellyfin` / `jf`                 | [Jellyfin](/jellyfin-integration) actions: `setup`, `login`, `logout`, `play`, `discovery`                           |
| `subminer dictionary <path>` / `dict`      | Build a [character dictionary](/character-dictionary) for a file or directory                                        |
| `subminer dictionary --candidates <path>`  | List AniList matches for that target                                                                                 |
| `subminer dictionary --select <id> <path>` | Pin an AniList ID for that target                                                                                    |
| `subminer texthooker`                      | Run only the texthooker server. `-o` opens it in your browser                                                        |
| `subminer logs -e`                         | Export a sanitized log ZIP and print its path                                                                        |
| `subminer config path` / `show`            | Print the config file path or its contents                                                                           |
| `subminer mpv status`                      | Exit 0 if the mpv socket is ready, 1 if not                                                                          |
| `subminer mpv socket`                      | Print the mpv socket path                                                                                            |
| `subminer mpv idle`                        | Start an idle mpv in the background with SubMiner's options                                                          |
| `subminer app` / `bin`                     | Pass arguments to the SubMiner app, e.g. `subminer app --stop`                                                       |

`stats cleanup` runs one mode at a time. `--lookback-days` must be at least 1. Without it, cleanup scans all history.

`generate-subs` options: `--download-model`, `--model <name>`, `--model-path <file>`, `--output <file>`, and `--audio-stream <index>` (an ffprobe stream index). `--model-path` cannot be combined with `--model` or `--download-model`.

A texthooker is a web page that shows the current subtitle as plain text, so browser extensions and other tools can read along.

## Video picker

With no file argument, `subminer` opens a picker for the current directory, or for `-d <dir>`. Add `-r` to include subdirectories.

- **fzf** (default) runs in the terminal. With `chafa` installed, it shows thumbnail previews.
- **rofi** (`-R`, Linux) opens a graphical menu with thumbnails.

Thumbnails come from your system thumbnail cache, or are generated with `ffmpegthumbnailer` or `ffmpeg`.

The launcher installs its rofi theme automatically. To use your own, set `SUBMINER_ROFI_THEME`:

```bash
SUBMINER_ROFI_THEME=/path/to/theme.rasi subminer -R
```

## Watch history

`subminer -H` lists the shows you have watched, most recent first. Add `-R` to use rofi. Pick a show, then choose:

- **Previous episode** or **Next episode**, moving into the neighboring season folder when needed
- **Replay last watched**
- **Browse episodes**, with a season menu first if the show has several season folders
- **Quit SubMiner**

When an episode ends, the menu comes back for the same show. Press `Escape` to leave.

History comes from the immersion stats database, which SubMiner fills during playback. Shows whose folders are not reachable, such as an unmounted network drive, are hidden.

## mpv options and profiles

The launcher starts mpv with these options:

```
--input-ipc-server=/tmp/subminer-socket
--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us
--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us
--sub-auto=fuzzy
--sub-file-paths=.;subs;subtitles
--sid=auto
--secondary-sid=auto
--sub-visibility=no
--secondary-sub-visibility=no
```

mpv's own subtitles are hidden because the overlay draws them. Add more options with `-a`, or load an mpv profile with `-p <name>` or `mpv.profile` in the config. No profile is loaded by default.

To launch mpv yourself with the same setup, put the options in a profile in `~/.config/mpv/mpv.conf`:

```ini
[subminer]
input-ipc-server=/tmp/subminer-socket
alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us
slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us
sub-auto=fuzzy
sub-file-paths=.;subs;subtitles
sid=auto
secondary-sid=auto
secondary-sub-visibility=no
```

Launches through `subminer` start the overlay automatically unless `mpv.autoStartSubMiner` is off. mpv started outside SubMiner does not start the overlay on its own.

## Sync between machines

`subminer sync <host>` merges immersion stats and watch history between two computers over SSH. Both end up with the combined sessions, totals, vocabulary, charts, and `-H` history. `<host>` is anything `ssh` accepts, such as `user@hostname` or an alias from your SSH config.

Both machines need the same SubMiner version. The remote only needs the app. The `subminer` command is optional there.

```bash
subminer sync macbook                    # two-way sync
subminer sync macbook --push             # send local data to macbook only
subminer sync macbook --pull             # bring macbook data here only
subminer sync macbook --check            # test SSH and the remote install, change nothing
subminer sync macbook --remote-cmd ~/Apps/SubMiner.AppImage   # SubMiner in a custom place on the remote
subminer sync --ui                       # open the sync window
```

Syncing only adds data. It never overwrites or double-counts, so you can run it as often as you like. `--push` and `--pull` do not delete anything on the receiving side.

Before a command-line sync, close SubMiner on both machines and stop the stats server with `subminer stats -s`, or pass `--force`. The sync window does not need this. It syncs while SubMiner and playback are running, and skips the session in progress until it finishes.

Transfers are compressed. When both machines have `rsync` (macOS and Linux), later syncs send only what changed. Windows machines use `scp`.

Known-word status from Anki does not sync. Each machine reads it from its own Anki collection.

<details>
<summary><b>More sync options</b></summary>

| Option              | Description                                                 |
| ------------------- | ----------------------------------------------------------- |
| `-f, --force`       | Skip the check that SubMiner is closed                      |
| `--db <file>`       | Use a different local stats database                        |
| `--json`            | Print progress as NDJSON                                    |
| `--snapshot <file>` | Write a snapshot of the local database, e.g. to copy by USB |
| `--merge <file>`    | Merge a snapshot file into the local database               |

A Windows remote needs the built-in **OpenSSH Server** enabled. SubMiner finds itself in the default install location there.

If the remote cannot find SubMiner, point `--remote-cmd` at the app or launcher, or link it as `SubMiner` somewhere on the remote `PATH`.

Received snapshots are cached in `sync-transfer-cache/` in the config directory to speed up later syncs. Deleting it is safe.

</details>

### Sync window

`subminer sync --ui`, or **Sync Stats & History** in the tray, opens a window where you can:

- Save devices, each with a direction (two-way, push, or pull), and run **Sync now** or **Test**.
- Turn on **Auto-sync** for a device. It syncs in the background every 60 minutes by default, including during playback.
- Watch progress and see what was merged on each machine.
- Create, merge, or delete database snapshots.

Saved devices live in `sync-hosts.json` in the config directory.

## Environment variables

| Variable                 | Use                                                                   |
| ------------------------ | --------------------------------------------------------------------- |
| `SUBMINER_BINARY_PATH`   | Path to the SubMiner app, if it is not in a standard install location |
| `SUBMINER_APPIMAGE_PATH` | Same, for an AppImage (Linux)                                         |
| `SUBMINER_ROFI_THEME`    | Path to a custom rofi theme                                           |

## Logging

The default log level is `warn`. Change it for one run with `--log-level`, or permanently with `logging.level` in the config. The app's `--dev` and `--debug` flags turn on developer mode. They do not change the log level.
