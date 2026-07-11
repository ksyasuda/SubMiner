# MPV Plugin

**What this is:** mpv is the video player SubMiner overlays subtitles on. The SubMiner mpv plugin is a small Lua script that runs _inside_ mpv and gives you in-player keybindings to control the SubMiner overlay (start/stop/toggle, skip intro, etc.) without leaving the player window.

**Who needs this page:** Most users never touch the plugin directly - SubMiner-managed launches (the app, the `subminer` launcher, or the Windows shortcut) inject the bundled plugin automatically for that session, so there is nothing to install into mpv's global `scripts` directory. Read on if you launch mpv from another tool and want SubMiner's in-player controls, or you want to script mpv against SubMiner.

The plugin ships as a modular Lua package under `plugin/subminer/` (entry point `main.lua`, which loads `init.lua` and sibling modules). Earlier releases shipped a single global `main.lua`; runtime loading replaces it.

## Runtime Loading

Launch mpv through the SubMiner app, the `subminer` launcher, or the packaged Windows SubMiner mpv shortcut. These paths pass mpv a bundled plugin path for that playback session only, leaving regular mpv playback untouched.

On Linux, the launcher-managed runtime plugin copy lives under the SubMiner data dir (`$XDG_DATA_HOME/SubMiner/plugin/subminer` by default, plus `/usr/local/share/SubMiner` or `/usr/share/SubMiner` for system installs). `subminer -u` and the tray updater keep that managed copy current. This is separate from mpv's global `scripts/` directory.

If setup detects an older global SubMiner plugin in mpv's `scripts` directory, use **Remove legacy mpv plugin** in first-run setup. The global plugin is not needed once runtime loading is available.

mpv must have IPC enabled for SubMiner to connect:

```ini
# ~/.config/mpv/mpv.conf
input-ipc-server=/tmp/subminer-socket
```

On Windows, use a named pipe instead:

```ini
input-ipc-server=\\.\pipe\subminer-socket
```

## Configuration (script-opts)

The plugin reads options from `script-opts` with the `subminer-` prefix (for example `--script-opts=subminer-backend=hyprland`). Managed launches inject these automatically from your SubMiner config; the shipped `subminer.conf` is intentionally empty so command-line opts always win.

| Option                                             | Default          | Description                                                                     |
| -------------------------------------------------- | ---------------- | -------------------------------------------------------------------------------- |
| `binary_path`                                      | `""`             | Path to the SubMiner binary; empty enables [auto-detection](#binary-auto-detection) |
| `socket_path`                                      | platform default | mpv IPC socket path (`/tmp/subminer-socket`, or `\\.\pipe\subminer-socket` on Windows) |
| `texthooker_enabled`                               | `no`             | Start the texthooker server with the overlay                                    |
| `texthooker_port`                                  | `5174`           | Texthooker server port                                                          |
| `backend`                                          | `auto`           | Window backend (`auto`, `hyprland`, `sway`, `x11`, `macos`)                     |
| `auto_start`                                       | `no`             | Start the overlay app on `file-loaded` (managed launches set this from `mpv.autoStartSubMiner`) |
| `auto_start_visible_overlay`                       | `no`             | Show the visible overlay on auto-start (from `auto_start_overlay` in config)    |
| `overlay_loading_osd`                              | `no`             | Show an OSD loading spinner while the overlay starts                            |
| `auto_start_pause_until_ready`                     | `yes`            | Keep mpv paused until the overlay reports tokenization-ready                    |
| `auto_start_pause_until_ready_timeout_seconds`     | `30`             | Timeout before resuming playback anyway                                         |
| `osd_messages`                                     | `yes`            | Show plugin OSD status messages                                                 |
| `log_level`                                        | `info`           | Plugin log verbosity                                                            |

## Keybindings

Most plugin actions use a `y` chord prefix - press `y`, then the second key (a "chord"):

| Chord           | Action                                 |
| --------------- | -------------------------------------- |
| `y-y`           | Open menu                              |
| `y-s`           | Start overlay                          |
| `y-S`           | Stop overlay                           |
| `y-t`           | Toggle visible overlay                 |
| `y-o`           | Open settings window                   |
| `y-r`           | Restart overlay                        |
| `y-c`           | Check status                           |
| `y-h`           | Open session help / keybinding modal   |
| `v`             | Toggle primary subtitle bar visibility |
| `TAB` (default) | Skip intro (AniSkip)                   |

The AniSkip key is **not** a `y` chord and is not bound by the plugin: the SubMiner app binds it over the mpv IPC socket while it is connected. It defaults to `TAB` and is configurable via `mpv.aniskipButtonKey`. When a custom key (other than `TAB` or `y-k`) is configured, the legacy `y-k` chord is also bound as a fallback. See [AniSkip Integration](/aniskip-integration) for setup and details.

The bare `v` binding is a forced mpv binding. It overrides mpv's default primary subtitle visibility toggle and routes the action to SubMiner's primary subtitle bar instead.

## Shared Shortcuts (Session Bindings)

The `y-*` chords above are built into the plugin. Everything else you configure under [`shortcuts.*`](/shortcuts) - plus any custom [`keybindings`](/configuration) and the stats toggle/mark-watched keys - is **injected into mpv at runtime**, so the same shortcut works both inside mpv and in the SubMiner overlay. You do not edit any mpv config to enable them.

How it works:

1. The SubMiner app compiles your configured shortcuts, custom keybindings, and stats keys into a normalized list and writes it to `session-bindings.json` in the SubMiner config directory.
2. On load, the plugin reads that file and registers each entry as a forced mpv key binding, translating each accelerator into the matching mpv key name.
3. When a binding fires, the plugin either runs a SubMiner action (by invoking the SubMiner binary with the corresponding CLI flag, e.g. `--mine-sentence`) or runs a raw mpv command, depending on what the shortcut maps to.

Because the bindings come from the same configuration the overlay uses, you maintain one set of shortcuts for both surfaces.

Live updates: changing a shortcut in the app rewrites `session-bindings.json` and sends the plugin a `subminer-reload-session-bindings` script message, so mpv re-registers the bindings immediately - no mpv restart required.

Notes:

- Accelerators are normalized per platform - `CommandOrControl` resolves to `Cmd` on macOS and `Ctrl` elsewhere.
- Multi-line actions (`copySubtitleMultiple`, `mineSentenceMultiple`) register temporary `1`–`9` digit follow-up bindings after the trigger key, with `Esc` to cancel.
- If two shortcuts compile to the same key, or an accelerator can't be mapped to an mpv key, the app logs a warning and skips that binding instead of registering a broken one.

## Menu

Press `y-y` to open an interactive menu (rendered with mpv's console selector):

```text
SubMiner:
1. Start overlay
2. Stop overlay
3. Toggle overlay
4. Open options
5. Restart overlay
6. Check status
7. Stats
```

Select an item by pressing its number.

## Binary Auto-Detection

When `binary_path` is empty, the plugin searches platform-specific locations:

**Linux:**

1. `~/.local/bin/SubMiner.AppImage`
2. `/opt/SubMiner/SubMiner.AppImage`
3. `/usr/local/bin/SubMiner` / `/usr/local/bin/subminer`
4. `/usr/bin/SubMiner` / `/usr/bin/subminer`

**macOS:**

1. `/Applications/SubMiner.app/Contents/MacOS/SubMiner`
2. `~/Applications/SubMiner.app/Contents/MacOS/SubMiner`

**Windows:**

A PowerShell system lookup runs first (running SubMiner process, registry App Paths, `Get-Command`), then static paths:

1. `%LOCALAPPDATA%\Programs\SubMiner\SubMiner.exe` (the default per-user install location)
2. `C:\Program Files\SubMiner\SubMiner.exe`
3. `C:\Program Files (x86)\SubMiner\SubMiner.exe`
4. `C:\SubMiner\SubMiner.exe`

On Windows the plugin also normalizes a Unix-style `socket_path` (`/tmp/subminer-socket`) to the named pipe `\\.\pipe\subminer-socket` at runtime.

## Backend Detection

When `backend=auto`, the plugin detects the window manager:

1. **macOS** - detected via platform or `OSTYPE`.
2. **Hyprland** - detected via `HYPRLAND_INSTANCE_SIGNATURE`.
3. **Sway** - detected via `SWAYSOCK`.
4. **X11** - detected via `XDG_SESSION_TYPE=x11` or `DISPLAY`.
5. **Fallback** - defaults to X11 with a warning.

::: tip Wayland is compositor-specific
Native Wayland support is only available for Hyprland and Sway. If you use a different Wayland compositor, auto-detection will fall back to X11 - both mpv and SubMiner must be running under Xwayland, and `xdotool` and `xwininfo` must be installed.
:::

## Script Messages

The plugin can be controlled from other mpv scripts or the mpv command line using script messages:

```
script-message subminer-start
script-message subminer-stop
script-message subminer-toggle
script-message subminer-menu
script-message subminer-options
script-message subminer-restart
script-message subminer-status
script-message subminer-autoplay-ready
script-message subminer-stats-toggle
script-message subminer-visible-overlay-shown
script-message subminer-visible-overlay-hidden
script-message subminer-managed-subtitles-loading
script-message subminer-overlay-loading-ready
script-message subminer-reload-session-bindings
```

The last five are primarily used by the SubMiner app to notify the plugin of overlay/loading state and to trigger session-binding reloads.

The AniSkip messages (`subminer-skip-intro`, `subminer-aniskip-refresh`) still exist, but they are handled by the SubMiner app over the IPC socket rather than by the plugin - see [AniSkip Integration](/aniskip-integration#triggering-from-mpv).

The `subminer-start` message accepts overrides:

```
script-message subminer-start backend=hyprland socket=/custom/path texthooker=no log-level=debug
```

`log-level` here controls only logging verbosity passed to SubMiner.
`--debug` is a separate app/dev-mode flag in the main CLI and should not be used here for logging.

## Lifecycle

For how the plugin's auto-start fits into the full launch sequence - including when the launcher starts the overlay instead of the plugin - see [Playback Startup Flow](./architecture#playback-startup-flow).

- **File loaded**: If `auto_start=yes`, the plugin starts the overlay.
- **Auto-start pause gate**: If `auto_start_visible_overlay=yes` and `auto_start_pause_until_ready=yes`, launcher starts mpv paused. On cold managed background startup, SubMiner opens the tray and visible overlay shell before tokenization warmups finish, then the plugin resumes playback after SubMiner reports tokenization-ready (with a 30-second timeout fallback).
- **Duplicate auto-start events**: Repeated `file-loaded` hooks while overlay is already running are ignored for auto-start triggers (prevents duplicate start attempts).
- **MPV shutdown**: The plugin clears its hover/OSD/gate state on shutdown; the overlay app notices the closed IPC socket and shuts itself down.
- **Texthooker**: When `texthooker_enabled=yes`, the plugin appends `--texthooker` to the overlay start command so the app starts the texthooker server alongside the overlay.

## Using with the `subminer` Wrapper

The `subminer` wrapper script handles mpv launch, socket setup, and overlay lifecycle automatically. You do not need the plugin if you always use the wrapper.

The plugin is useful when you:

- Launch mpv from other tools (file managers, media centers).
- Want on-demand overlay control without the wrapper.
- Use mpv's built-in file browser or playlist features.

You can install both - the plugin provides chord keybindings for convenience, while the wrapper handles the full lifecycle.
