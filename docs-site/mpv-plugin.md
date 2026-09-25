# MPV plugin

The SubMiner mpv plugin is a Lua script that runs inside mpv. It adds in-player keys to start, stop, and toggle the overlay, and it runs your SubMiner shortcuts from inside mpv.

## Setup

You usually do not install anything. Every SubMiner-managed launch (the app, the `subminer` launcher, and the Windows SubMiner mpv shortcut) loads the bundled plugin for that session only. Regular mpv playback is not affected.

On Linux, the launcher's copy lives in `$XDG_DATA_HOME/SubMiner/plugin/subminer` (default `~/.local/share/SubMiner/plugin/subminer`), or under `/usr/local/share/SubMiner` or `/usr/share/SubMiner` for system installs. `subminer -u` and the tray updater keep it current.

To use the plugin when mpv is started by another program, load its `main.lua` and enable IPC:

```bash
mpv --script="$HOME/.local/share/SubMiner/plugin/subminer/main.lua" \
  --input-ipc-server=/tmp/subminer-socket video.mkv
```

To enable IPC for every mpv session, add it to `mpv.conf`:

```ini
input-ipc-server=/tmp/subminer-socket
```

On Windows, use a named pipe:

```ini
input-ipc-server=\\.\pipe\subminer-socket
```

If first-run setup finds an old SubMiner plugin in mpv's global `scripts` directory, click **Remove legacy mpv plugin**. It is no longer needed.

## Keybindings

Most plugin keys are chords: press `y`, then the second key.

| Key   | Action                                 |
| ----- | -------------------------------------- |
| `y-y` | Open the SubMiner menu                 |
| `y-s` | Start the overlay                      |
| `y-S` | Stop the overlay                       |
| `y-t` | Toggle the visible overlay             |
| `y-o` | Open the settings window               |
| `y-r` | Restart the overlay                    |
| `y-c` | Check status                           |
| `y-h` | Open the session help modal            |
| `v`   | Toggle SubMiner's primary subtitle bar |

`v` replaces mpv's own subtitle visibility toggle.

The skip-intro key (`TAB` by default) comes from the SubMiner app, not the plugin. See [AniSkip integration](/aniskip-integration).

The `y-y` menu lists Start overlay, Stop overlay, Toggle overlay, Open options, Restart overlay, Check status, and Stats. Press an item's number to run it. Stats only reminds you to press `` ` `` in the overlay.

## Your shortcuts in mpv

Everything you set under [`shortcuts`](/shortcuts), your custom `keybindings`, and the stats keys also work while mpv has focus. SubMiner writes them to `session-bindings.json` in its config directory, and the plugin registers them as mpv keys. When you change a shortcut, mpv picks it up immediately.

`CommandOrControl` becomes `Cmd` on macOS and `Ctrl` elsewhere. Multi-line copy and mine shortcuts wait for a digit key `1` to `9`, and `Esc` cancels. If two shortcuts map to the same key, or a key has no mpv equivalent, SubMiner logs a warning and skips it.

## Script options

The plugin reads `script-opts` with the `subminer-` prefix, for example `--script-opts=subminer-backend=hyprland`. Managed launches set these from your SubMiner config, so edit the config instead. The shipped `plugin/subminer.conf` is empty on purpose, so it never overrides those values.

| Option                                         | Default          | SubMiner config key          | What it does                                                          |
| ---------------------------------------------- | ---------------- | ---------------------------- | --------------------------------------------------------------------- |
| `binary_path`                                  | `""`             | `mpv.subminerBinaryPath`     | SubMiner binary. Empty uses [auto-detection](#binary-auto-detection). |
| `socket_path`                                  | platform default | `mpv.socketPath`             | mpv IPC socket                                                        |
| `backend`                                      | `auto`           | `mpv.backend`                | Window backend: `auto`, `hyprland`, `sway`, `x11`, `macos`            |
| `auto_start`                                   | `no`             | `mpv.autoStartSubMiner`      | Start SubMiner when a file loads                                      |
| `auto_start_visible_overlay`                   | `no`             | `auto_start_overlay`         | Show the overlay when auto-starting                                   |
| `auto_start_pause_until_ready`                 | `yes`            | `mpv.pauseUntilOverlayReady` | Keep mpv paused until subtitles are ready                             |
| `auto_start_pause_until_ready_timeout_seconds` | `30`             |                              | Resume anyway after this many seconds                                 |
| `overlay_loading_osd`                          | `no`             |                              | Show a loading message while the overlay starts                       |
| `texthooker_enabled`                           | `no`             |                              | Start the texthooker with the overlay                                 |
| `texthooker_port`                              | `5174`           |                              | Texthooker port                                                       |
| `osd_messages`                                 | `yes`            |                              | Show plugin status messages in mpv                                    |
| `log_level`                                    | `info`           |                              | Plugin log level                                                      |

Without script options, `socket_path` is `/tmp/subminer-socket`, or `\\.\pipe\subminer-socket` on Windows. On Windows, the plugin also rewrites `/tmp/subminer-socket` to the named pipe.

The table's defaults are the plugin's own. Managed launches override them from your config; see [Configuration](/configuration#mpv-launcher).

## Binary auto-detection

With `binary_path` empty, the plugin looks in these places:

| Platform | Locations                                                                                                                                                                                                  |
| -------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Linux    | `~/.local/bin/SubMiner.AppImage`, `/opt/SubMiner/SubMiner.AppImage`, `/usr/local/bin/SubMiner` or `subminer`, `/usr/bin/SubMiner` or `subminer`                                                            |
| macOS    | `/Applications/SubMiner.app`, `~/Applications/SubMiner.app`                                                                                                                                                |
| Windows  | A running SubMiner process, the App Paths registry entry, `SubMiner.exe` on `PATH`, then `%LOCALAPPDATA%\Programs\SubMiner`, `C:\Program Files\SubMiner`, `C:\Program Files (x86)\SubMiner`, `C:\SubMiner` |

## Backend detection

With `backend=auto`, the plugin picks the first match:

1. macOS
2. Hyprland (`HYPRLAND_INSTANCE_SIGNATURE` is set)
3. Sway (`SWAYSOCK` is set)
4. X11 (`XDG_SESSION_TYPE=x11` or `DISPLAY` is set)
5. Otherwise X11, with a warning

Native Wayland support covers only Hyprland and Sway. On other Wayland compositors, run both mpv and SubMiner under Xwayland and install `xdotool` and `xwininfo`.

## Script messages

Other mpv scripts, `input.conf`, or the mpv console can control the plugin:

```text
script-message subminer-start
script-message subminer-stop
script-message subminer-toggle
script-message subminer-menu
script-message subminer-options
script-message subminer-restart
script-message subminer-status
```

`subminer-start` accepts overrides:

```text
script-message subminer-start backend=hyprland socket=/custom/path texthooker=no log-level=debug
```

`log-level` sets SubMiner's log verbosity. Do not use `--debug` for this; it turns on the app's dev mode.

The plugin also handles messages the SubMiner app sends it (`subminer-autoplay-ready`, `subminer-visible-overlay-shown`, `subminer-visible-overlay-hidden`, `subminer-managed-subtitles-loading`, `subminer-overlay-loading-ready`, `subminer-reload-session-bindings`). You do not need to send these yourself. The AniSkip messages are listed on the [AniSkip page](/aniskip-integration#triggering-from-mpv).

## Auto-start behavior

- With `auto_start=yes`, the plugin starts SubMiner on each file load. Repeated loads while SubMiner is running do not start it again.
- With `auto_start_visible_overlay=yes` and `auto_start_pause_until_ready=yes`, mpv stays paused until SubMiner reports that subtitles are ready, or until the timeout passes.
- With `texthooker_enabled=yes`, the texthooker starts with the overlay.
- When mpv quits, SubMiner sees the closed socket and shuts down its overlay.
