# Usage

## Quick Start

Play a video with SubMiner:

```bash
subminer video.mkv
```

On **Windows**, use the **SubMiner mpv** shortcut created during first-run setup - double-click it, or drag a video file onto it.

That's the simplest way to get started. The `subminer` launcher handles mpv, the IPC socket, and the overlay automatically.

> [!IMPORTANT]
> SubMiner requires the bundled Yomitan instance to have at least one dictionary imported for lookups to work.
> See [Yomitan setup](#yomitan-setup) for details.

::: tip Anki card enrichment
If you want sentence, audio, and screenshot fields on your Anki cards, add this to your config:

```jsonc
{
  "ankiConnect": {
    "enabled": true,
    "deck": "Mining",
    "fields": {
      "sentence": "Sentence",
      "audio": "SentenceAudio",
      "image": "Picture",
    },
  },
}
```

Field names must match your Anki note type exactly (case-sensitive). See [Anki Integration](/anki-integration) for the full reference.
:::

## How It Works

When you launch SubMiner, it wires up mpv and the overlay for you:

1. SubMiner starts the overlay app in the background
2. mpv runs with an **IPC socket** at `/tmp/subminer-socket` - a small local channel two programs use to talk to each other, so the overlay can ask mpv what subtitle is on screen right now
3. The overlay connects and subscribes to subtitle changes

From there, subtitles render as interactive, hoverable word spans and you mine cards directly from the overlay. For the overlay anatomy and the full mining loop - word lookup, card creation, annotations - see [Mining Workflow](/mining-workflow).

### Ways to Launch

| Approach                            | Use when                                                                                                                               | How                                                                   |
| ----------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------- |
| **`subminer` launcher**             | You want SubMiner to handle everything - launch mpv, set up the socket, start the overlay. **Recommended for most users.**             | `subminer video.mkv`                                                  |
| **SubMiner mpv shortcut** (Windows) | The recommended Windows entry point. Created during first-run setup, launches mpv with SubMiner's defaults.                            | Double-click, drag a file onto it, or run `SubMiner.exe --launch-mpv` |
| **mpv plugin** (all platforms)      | Bundled and injected at runtime. Provides `y` chord keybindings for controlling the overlay from within mpv. No manual install needed. | Automatic when using the launcher or shortcut                         |

The mpv plugin is always available - it's bundled with SubMiner and injected at runtime. On Linux, normal `subminer` playback auto-installs the launcher-managed runtime plugin copy from the bundled app if that managed copy is missing, so no separate plugin install is needed for standard launcher usage. If you launch mpv yourself (without the launcher), pass `--input-ipc-server=/tmp/subminer-socket` in your mpv config for the overlay to connect.

## Live Config Reload

While SubMiner is running, it watches your active config file and applies safe updates automatically.

Live-updated settings:

- `subtitleStyle`
- `keybindings`
- `shortcuts`
- `secondarySub.defaultMode`
- `ankiConnect.ai`

Invalid config edits are rejected; SubMiner keeps the previous valid runtime config and shows an error notification.
For restart-required sections, SubMiner shows a restart-needed notification.

## Commands

On Windows, replace `SubMiner.AppImage` with `SubMiner.exe` in the direct packaged-app examples below.

```bash
# Browse and play videos
subminer                          # Current directory (uses fzf)
subminer -R                       # Use rofi instead of fzf
subminer -d ~/Videos              # Specific directory
subminer -r -d ~/Anime            # Recursive search
subminer video.mkv                # Play specific file (overlay auto-starts)
subminer --start video.mkv        # Explicit overlay start (use when auto_start=no in config)
subminer -S video.mkv             # Same as above via --start-overlay
subminer https://youtu.be/...     # Play a YouTube URL
subminer ytsearch:"jp news"       # Play first YouTube search result
subminer --setup                  # Open first-run setup popup
subminer --version                # Print installed SubMiner version
subminer -v                       # Same as above
subminer --log-level debug video.mkv # Enable verbose logs for launch/debugging
subminer --log-level warn video.mkv  # Set logging level explicitly
subminer --args '--fs=opengl-hq --ytdl-format=bestvideo*+bestaudio/best' video.mkv  # Pass extra mpv args

# Options
subminer -T video.mkv             # Disable texthooker server
subminer -b x11 video.mkv         # Force X11 backend
subminer video.mkv                # No mpv profile passed by default
subminer -p gpu-hq video.mkv      # Use a specific mpv profile
subminer jellyfin                 # Open Jellyfin setup window (subcommand form)
subminer jellyfin -l --server http://127.0.0.1:8096 --username me --password 'secret'
subminer jellyfin --logout        # Clear stored Jellyfin token/session data
subminer jellyfin -p              # Interactive Jellyfin library/item picker + playback
subminer jellyfin -d              # Jellyfin cast-discovery mode (background tray app)
subminer app --stop               # Stop background app (including Jellyfin cast broadcast)
subminer doctor                   # Dependency + config + socket diagnostics
subminer logs -e                  # Export a sanitized log ZIP and print its path
subminer config path              # Print active config path
subminer config show              # Print active config contents
subminer mpv socket               # Print active mpv socket path
subminer mpv status               # Exit 0 if socket is ready, else exit 1
subminer mpv idle                 # Launch detached idle mpv with SubMiner defaults
subminer dictionary /path/to/file-or-directory  # Generate character dictionary ZIP from target (manual Yomitan import)
subminer dictionary --candidates /path/to/file.mkv
subminer dictionary --select 21355 /path/to/file.mkv
subminer texthooker               # Launch texthooker-only mode
subminer texthooker -o            # Launch texthooker and open it in your browser
subminer app --anilist-setup      # Pass args directly to SubMiner binary (example: AniList login flow)

# Direct packaged app control
SubMiner.AppImage --background             # Start in background (tray + IPC wait, minimal logs)
SubMiner.AppImage --start --texthooker   # Start overlay with texthooker
SubMiner.AppImage --texthooker           # Launch texthooker only (no overlay window)
SubMiner.AppImage --texthooker --open-browser  # Launch texthooker and open browser
SubMiner.AppImage --setup                  # Open first-run setup popup
SubMiner.AppImage --stop                  # Stop overlay
SubMiner.AppImage --start --toggle        # Start MPV IPC + toggle visibility
SubMiner.AppImage --show-visible-overlay              # Force show visible overlay
SubMiner.AppImage --hide-visible-overlay              # Force hide visible overlay
SubMiner.AppImage --toggle-primary-subtitle-bar       # Toggle primary subtitle bar visibility
SubMiner.AppImage --start --dev                         # Enable app/dev mode only
SubMiner.AppImage --start --debug                       # Alias for --dev
SubMiner.AppImage --start --log-level debug             # Force verbose logging without app/dev mode
SubMiner.AppImage --playback-feedback "your feedback"   # Route playback feedback through the configured feedback surface
SubMiner.AppImage --yomitan               # Open Yomitan settings
SubMiner.AppImage --settings              # Open SubMiner settings window
SubMiner.AppImage --jellyfin              # Open Jellyfin setup window
SubMiner.AppImage --jellyfin-login --jellyfin-server http://127.0.0.1:8096 --jellyfin-username me --jellyfin-password 'secret'
SubMiner.AppImage --jellyfin-logout       # Clear stored Jellyfin token/session data
SubMiner.AppImage --jellyfin-libraries
SubMiner.AppImage --jellyfin-items --jellyfin-library-id LIBRARY_ID --jellyfin-search anime --jellyfin-limit 20
SubMiner.AppImage --jellyfin-play --jellyfin-item-id ITEM_ID --jellyfin-audio-stream-index 1 --jellyfin-subtitle-stream-index 2  # Requires connected mpv IPC (--start)
SubMiner.AppImage --jellyfin-remote-announce  # Force cast-target capability announce + visibility check
SubMiner.AppImage --dictionary             # Generate character dictionary ZIP for current anime
SubMiner.AppImage --dictionary-candidates  # List AniList candidates for current character dictionary series
SubMiner.AppImage --dictionary-select --dictionary-anilist-id 21355  # Pin correct AniList media for series
SubMiner.AppImage --help                  # Show all options
```

The tray menu includes `Export Logs`, which creates the same sanitized local-date log ZIP as `subminer logs -e` and shows the archive path when complete. Export sanitization masks common PII and secrets, including home-directory usernames, IP addresses, emails, auth/cookie headers, yt-dlp cookie arguments, URL credentials, token/key/password fields, and signed YouTube media URL query strings. The exported copy is sanitized; source log files remain unredacted on disk.

Once Jellyfin is configured, the tray menu includes `Jellyfin Discovery` for starting or stopping cast discovery in the current app session without changing config.

### Logging and App Mode

- `--log-level` controls logger verbosity.
- `--dev` and `--debug` are app/dev-mode switches; they are not log-level aliases.
- `--background` defaults to quieter logging (`warn`) unless `--log-level` is set.
- `--background` launched from a terminal detaches and returns the prompt; stop it with tray Quit or `SubMiner.AppImage --stop` (`SubMiner.exe --stop` on Windows).
- Linux desktop launcher starts SubMiner with `--background` by default (via electron-builder `linux.executableArgs`).
- On Hyprland and other Wayland compositors, the tray icon appears only when your panel provides a StatusNotifier/AppIndicator tray host.
- On Linux, the app now defaults `safeStorage` to `gnome-libsecret` for encrypted token persistence.
  Launcher pass-through commands also support `--password-store=<backend>` and forward it to the app when present.
  Override with e.g. `--password-store=basic_text`.
- Use both when needed, for example `SubMiner.AppImage --start --dev --log-level debug` (or `SubMiner.exe --start --dev --log-level debug` on Windows).
- `--playback-feedback <text>` (also `--playback-feedback=<text>`) sends a non-empty text string through the playback-feedback route used for recording/playback prompts. For example: `SubMiner.AppImage --playback-feedback "your feedback"`.

### Windows mpv Shortcut

First-run setup creates the config file, then requires Yomitan dictionaries before it can finish.

If you enabled the optional Windows shortcut during install, SubMiner creates a `SubMiner mpv` shortcut in the Start menu and/or on the desktop. On Windows, that shortcut is the recommended way to launch local files with SubMiner because it starts `mpv.exe` with the right defaults directly.
After setup completes, the shortcut is the normal Windows playback entry point.

You can use it three ways:

- Double-click `SubMiner mpv` to open `mpv` with SubMiner's default socket/subtitle args.
- Drag a video file onto `SubMiner mpv` to launch that file with the same defaults.
- Run it directly from Command Prompt or PowerShell with `--launch-mpv`.

```powershell
& "C:\Program Files\SubMiner\SubMiner.exe" --launch-mpv
& "C:\Program Files\SubMiner\SubMiner.exe" --launch-mpv "C:\Videos\episode 01.mkv"
```

This flow requires `mpv.exe` to be discoverable. Leave `mpv.executablePath` blank to auto-discover from `PATH`, or set it to the full `mpv.exe` path if mpv is installed elsewhere. `SUBMINER_MPV_PATH` is still honored as a fallback.

### Launcher Subcommands

- `subminer jellyfin` / `subminer jf`: Jellyfin-focused workflow aliases.
- `subminer doctor`: health checks for core dependencies and runtime paths.
- `subminer settings`: open the SubMiner settings window (also `subminer --settings`).
- `subminer logs -e`: export a sanitized ZIP of today's local-date logs, or the most recent logs when no current-day log exists. The exported copy masks common PII and secrets; on-disk logs are unchanged.
- `subminer config`: config file helpers (`path`, `show`).
- `subminer mpv`: mpv helpers (`status`, `socket`, `idle`).
- `subminer dictionary <path>`: generates a Yomitan-importable character dictionary ZIP from a file/directory target.
- Use `subminer dictionary --candidates <path>` and `subminer dictionary --select <id> <path>` to correct AniList character-dictionary matches for a whole series.
- `subminer texthooker`: texthooker-only shortcut (same behavior as `--texthooker`). A _texthooker_ is a web page that displays the current subtitle line as selectable text, so browser-based dictionary extensions and other tools can read along with playback.
- `subminer app` / `subminer bin`: direct passthrough to the SubMiner binary/AppImage.
- Subcommand help pages are available (for example `subminer jellyfin -h`).

### First-Run Setup

Setup popup appears on first launch, or when setup has not been completed.

You can also open it manually:

```bash
subminer --setup
SubMiner.AppImage --setup
```

Setup flow:

- config file: create the default config directory and prefer `config.jsonc`
- legacy plugin cleanup: remove detected older global SubMiner mpv plugin files if present (the bundled plugin is injected at runtime automatically)
- Yomitan shortcut: open bundled Yomitan settings directly from the setup window
- dictionary check: ensure at least one bundled Yomitan dictionary is available, unless an external Yomitan profile is configured
- Windows: optionally create or remove `SubMiner mpv` Start Menu/Desktop shortcuts (`SubMiner.exe --launch-mpv`)
- Windows: optionally set `mpv.executablePath` if `mpv.exe` is not on `PATH`
- refresh: re-check dictionary state without restarting
- `Finish setup` stays disabled until the config and dictionary gates are satisfied
- finish action writes setup completion state and suppresses future auto-open prompts

AniList character dictionary auto-sync (optional):

- Enable with `subtitleStyle.nameMatchEnabled=true` in config or **Name Match Enabled** in Settings.
- SubMiner syncs the currently watched AniList media into a per-media snapshot, then rebuilds one merged `SubMiner Character Dictionary` from the most recently used snapshots.
- Rotation limit defaults to 3 recent media snapshots in that merged dictionary (`maxLoaded`).

Use subcommands for Jellyfin workflows (`subminer jellyfin ...`).
Top-level launcher flags like `--jellyfin-*` are intentionally rejected.

### MPV Profile Example (mpv.conf)

`subminer` passes the following MPV options directly on launch by default:

- `--input-ipc-server=/tmp/subminer-socket` (or your configured socket path)
- `--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us`
- `--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us`
- `--sub-auto=fuzzy`
- `--sub-file-paths=.;subs;subtitles`
- `--sid=auto`
- `--secondary-sid=auto`
- `--secondary-sub-visibility=no`

You can append additional MPV arguments with launcher `-a/--args`, for example `--args "--ao=alsa --volume=80"`.

You can define a matching profile in `~/.config/mpv/mpv.conf` for consistency when launching `mpv` manually or from other tools. The Windows `SubMiner.exe --launch-mpv` shortcut path uses equivalent args directly, but skips the extra current-directory subtitle scan to avoid duplicate sidecar detection when you drag a video onto the shortcut; the optional profile remains useful for manual mpv launches. The `subminer` wrapper passes no mpv profile by default; set one with `subminer -p <profile> ...` or with `mpv.profile` in your config (for example `"profile": "subminer"` to use the `[subminer]` profile below):

```ini
[subminer]
# IPC socket (must match SubMiner config)
input-ipc-server=/tmp/subminer-socket

# Prefer JP/EN audio + subtitle language variants
alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us
slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us

# Auto-load external subtitles
sub-auto=fuzzy
sub-file-paths=.;subs;subtitles

# Select primary + secondary subtitle tracks automatically
sid=auto
secondary-sid=auto
secondary-sub-visibility=no
```

### Yomitan setup

SubMiner includes a bundled Yomitan extension for overlay word lookup. This bundled extension is separate from any Yomitan browser extension you may have installed.

For SubMiner overlay lookups to work, open Yomitan settings (`subminer app --yomitan` or `SubMiner.AppImage --yomitan`) and import at least one dictionary in the bundled Yomitan instance.

If you also use Yomitan in a browser, configure that browser profile separately; it does not inherit dictionaries or settings from the bundled instance.

### YouTube Playback

`subminer` accepts direct URLs (for example, YouTube links) and `ytsearch:` targets.
For YouTube playback, SubMiner resolves subtitle selection during startup while mpv is paused: it auto-selects the default primary subtitle track plus a best-effort secondary track, then resumes when primary subtitles are ready.

Notes:

- Install `yt-dlp` so mpv can resolve YouTube streams and subtitle tracks reliably.
- For YouTube URLs, startup no longer requires opening the picker first; SubMiner loads subtitles and keeps the overlay available for retries.
- Press `Ctrl+Alt+C` during active YouTube playback to open the manual YouTube subtitle picker and retry track selection.
- For YouTube URLs, `subminer` probes available YouTube subtitle tracks, reuses existing authoritative tracks when available, and downloads only missing sides.
- Native mpv secondary subtitle rendering stays hidden so the overlay remains the visible secondary subtitle surface.
- Primary subtitle target languages come from `youtube.primarySubLanguages` (defaults to `["ja","jpn"]`).
- Secondary target languages come from `secondarySub.secondarySubLanguages` (empty by default; when empty, no language-based secondary track is auto-selected, though mpv's `--slang` list above still prefers English variants). When multiple matching secondary tracks exist, SubMiner prefers a non-Signs/Songs track.
- Configure defaults in `$XDG_CONFIG_HOME/SubMiner/config.jsonc` (or `~/.config/SubMiner/config.jsonc`) under `youtube` and `secondarySub`.

For local video files, SubMiner uses the same config-driven language priorities to auto-select the primary and secondary subtitle tracks from internal and external subtitle sources.

## Controller Support

SubMiner supports gamepad/controller input for couch-friendly usage via the Chrome Gamepad API. Controller input drives the overlay while keyboard-only mode is enabled.

### Getting Started

1. Connect a controller before or after launching SubMiner.
2. Set `controller.enabled` to `true` in your config.
3. Press `Alt+C` in the overlay by default to pick the controller you want to save and remap any action inline.
4. Enable keyboard-only mode - press `Y` on the controller (default binding) or use the overlay keybinding.
5. Click the binding badge, edit pencil, or `Learn` on the overlay action you want, then press the matching button, trigger, or stick direction on the controller.
6. Use the left stick to navigate subtitle tokens and scroll the popup; use the right stick vertically for popup page jumps.
7. Press `A` to look up the selected word, `X` to mine a card, `B` to close the popup.

By default SubMiner uses the first connected controller after controller support is enabled. `Alt+C` opens the controller config modal, where you can save the preferred controller and remap bindings inline per controller. The reset button beside each edit pencil restores that binding to its built-in default for the selected controller. `Alt+Shift+C` opens the live debug modal with raw axes/button values for non-standard pads. Both modals stay closed while `controller.enabled` is false, and both shortcuts can be changed through `shortcuts.openControllerSelect` and `shortcuts.openControllerDebug`.

### Default Button Mapping

| Button                  | Action                                  |
| ----------------------- | --------------------------------------- |
| `A` (South)             | Toggle lookup                           |
| `B` (East)              | Close lookup                            |
| `Y` (North)             | Toggle keyboard-only mode               |
| `X` (West)              | Mine card                               |
| `L1`                    | Play current Yomitan audio              |
| `R1`                    | Next Yomitan audio track                |
| `L3` (left stick press) | Toggle mpv pause                        |
| `Select` / `Minus`      | Quit mpv                                |
| `L2` / `R2`             | Unbound (available for custom bindings) |

### Analog Controls

| Input                 | Action                                        |
| --------------------- | --------------------------------------------- |
| Left stick horizontal | Move token selection left/right               |
| Left stick vertical   | Scroll Yomitan popup                          |
| Right stick vertical  | Jump through Yomitan popup                    |
| D-pad                 | Fallback for stick navigation when configured |

Learn mode ignores already-held inputs and waits for the next fresh button press or axis direction, which avoids accidental captures when you open the modal mid-input.

All button and axis mappings are configurable under the `controller` config block. Learned remaps are saved under `controller.profiles` for the selected controller id. See [Configuration - Controller Support](/configuration#controller-support) for the full options.

## Keybindings

See [Keyboard Shortcuts](/shortcuts) for the full reference, including mining shortcuts, overlay controls, and customization.

**Global shortcuts** (work system-wide):

| Keybind       | Action                 |
| ------------- | ---------------------- |
| `Alt+Shift+O` | Toggle visible overlay |
| `Alt+Shift+Y` | Open Yomitan settings  |

`Alt+Shift+Y` is fixed and not configurable. All other shortcuts can be changed under `shortcuts` in your config.

Useful overlay-local default keybinding: `Ctrl+Alt+P` opens the playlist browser for the current video's parent directory and the live mpv queue so you can append, reorder, remove, or jump between episodes without leaving playback.

Press `V` to hide or restore the primary SubMiner subtitle bar. The bundled mpv plugin also binds bare `v` to the same action (injected at runtime).

`Ctrl/Cmd+/` opens the session help modal with the current overlay and mpv keybindings. The same help view is also available through the `y-h` chord in mpv.

Hovering over subtitle text pauses mpv by default; leaving resumes it. Yomitan popups also pause playback by default. Set `subtitleStyle.autoPauseVideoOnHover: false` or `subtitleStyle.autoPauseVideoOnYomitanPopup: false` to disable either behavior.

### Drag-and-Drop

- Drop video files onto the overlay to replace current playback.
- Hold `Shift` while dropping to append to the playlist instead.

Next: [Mining Workflow](/mining-workflow) - word lookup, card creation, and the full mining loop.
