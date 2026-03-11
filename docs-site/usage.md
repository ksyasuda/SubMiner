# Usage

> [!IMPORTANT]
> SubMiner requires the bundled Yomitan instance to have at least one dictionary imported for lookups to work.
> See [Yomitan setup](#yomitan-setup) for details.

## How It Works

1. SubMiner starts the overlay app in the background
2. MPV runs with an IPC socket at `/tmp/subminer-socket`
3. The overlay connects and subscribes to subtitle changes
4. Subtitles are tokenized with Yomitan's internal parser
5. Words are displayed as interactive spans in the overlay
6. Hovering or clicking a word triggers Yomitan popup for dictionary lookup
7. Optional [subtitle annotations](/subtitle-annotations) (N+1, character-name, frequency, JLPT) highlight useful cues in real time

There are two ways to use SubMiner:

| Approach | Use when | How |
| -------- | -------- | --- |
| **`subminer` script** | You want SubMiner to handle everything — launch mpv, set up the socket, start the overlay. The simplest path. | `subminer video.mkv` |
| **MPV plugin** | You launch mpv yourself or from another tool (file manager, Jellyfin, etc.). Requires `--input-ipc-server=/tmp/subminer-socket` in your mpv config. | Use `y` chord keybindings inside mpv |

You can use both — the plugin provides in-player controls, while the `subminer` script is convenient for direct playback. The `subminer` script runs directly via shebang (no `bun run` needed).

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
subminer video.mkv                # Play specific file (default plugin config auto-starts visible overlay)
subminer --start video.mkv        # Optional explicit overlay start (use when plugin auto_start=no)
subminer -S video.mkv             # Same as above via --start-overlay
subminer https://youtu.be/...     # Play a YouTube URL
subminer ytsearch:"jp news"       # Play first YouTube search result
subminer --setup                  # Open first-run setup popup
subminer --log-level debug video.mkv # Enable verbose logs for launch/debugging
subminer --log-level warn video.mkv  # Set logging level explicitly

# Options
subminer -T video.mkv             # Disable texthooker server
subminer -b x11 video.mkv         # Force X11 backend
subminer video.mkv                # Uses mpv profile "subminer" by default
subminer -p gpu-hq video.mkv      # Override mpv profile
subminer jellyfin                 # Open Jellyfin setup window (subcommand form)
subminer jellyfin -l --server http://127.0.0.1:8096 --username me --password 'secret'
subminer jellyfin --logout        # Clear stored Jellyfin token/session data
subminer jellyfin -p              # Interactive Jellyfin library/item picker + playback
subminer jellyfin -d              # Jellyfin cast-discovery mode (background tray app)
subminer app --stop               # Stop background app (including Jellyfin cast broadcast)
subminer doctor                   # Dependency + config + socket diagnostics
subminer config path              # Print active config path
subminer config show              # Print active config contents
subminer mpv socket               # Print active mpv socket path
subminer mpv status               # Exit 0 if socket is ready, else exit 1
subminer mpv idle                 # Launch detached idle mpv with SubMiner defaults
subminer dictionary /path/to/file-or-directory  # Generate character dictionary ZIP from target (manual Yomitan import)
subminer texthooker               # Launch texthooker-only mode
subminer app --anilist            # Pass args directly to SubMiner binary (example: AniList login flow)
subminer yt -o ~/subs https://youtu.be/...  # YouTube subcommand: output directory shortcut
subminer yt --keep-temp --whisper-bin /path/to/whisper-cli --whisper-model /path/to/model.bin --whisper-vad-model /path/to/ggml-vad.bin https://youtu.be/...  # Keep generated subtitle workspace for debugging

# Direct packaged app control
SubMiner.AppImage --background             # Start in background (tray + IPC wait, minimal logs)
SubMiner.AppImage --start --texthooker   # Start overlay with texthooker
SubMiner.AppImage --texthooker           # Launch texthooker only (no overlay window)
SubMiner.AppImage --setup                  # Open first-run setup popup
SubMiner.AppImage --stop                  # Stop overlay
SubMiner.AppImage --start --toggle        # Start MPV IPC + toggle visibility
SubMiner.AppImage --show-visible-overlay              # Force show visible overlay
SubMiner.AppImage --hide-visible-overlay              # Force hide visible overlay
SubMiner.AppImage --start --dev                         # Enable app/dev mode only
SubMiner.AppImage --start --debug                       # Alias for --dev
SubMiner.AppImage --start --log-level debug             # Force verbose logging without app/dev mode
SubMiner.AppImage --settings              # Open Yomitan settings
SubMiner.AppImage --jellyfin              # Open Jellyfin setup window
SubMiner.AppImage --jellyfin-login --jellyfin-server http://127.0.0.1:8096 --jellyfin-username me --jellyfin-password 'secret'
SubMiner.AppImage --jellyfin-logout       # Clear stored Jellyfin token/session data
SubMiner.AppImage --jellyfin-libraries
SubMiner.AppImage --jellyfin-items --jellyfin-library-id LIBRARY_ID --jellyfin-search anime --jellyfin-limit 20
SubMiner.AppImage --jellyfin-play --jellyfin-item-id ITEM_ID --jellyfin-audio-stream-index 1 --jellyfin-subtitle-stream-index 2  # Requires connected mpv IPC (--start or plugin workflow)
SubMiner.AppImage --jellyfin-remote-announce  # Force cast-target capability announce + visibility check
SubMiner.AppImage --dictionary             # Generate character dictionary ZIP for current anime
SubMiner.AppImage --help                  # Show all options
```

### Logging and App Mode

- `--log-level` controls logger verbosity.
- `--dev` and `--debug` are app/dev-mode switches; they are not log-level aliases.
- `--background` defaults to quieter logging (`warn`) unless `--log-level` is set.
- `--background` launched from a terminal detaches and returns the prompt; stop it with tray Quit or `SubMiner.AppImage --stop` (`SubMiner.exe --stop` on Windows).
- Linux desktop launcher starts SubMiner with `--background` by default (via electron-builder `linux.executableArgs`).
- On Linux, the app now defaults `safeStorage` to `gnome-libsecret` for encrypted token persistence.
  Launcher pass-through commands also support `--password-store=<backend>` and forward it to the app when present.
  Override with e.g. `--password-store=basic_text`.
- Use both when needed, for example `SubMiner.AppImage --start --dev --log-level debug` (or `SubMiner.exe --start --dev --log-level debug` on Windows).

### Windows mpv Shortcut

If you enabled the optional Windows shortcut during install, SubMiner creates a `SubMiner mpv` shortcut in the Start menu and/or on the desktop. It runs `SubMiner.exe --launch-mpv`, which starts `mpv.exe` with SubMiner's `subminer` profile.

You can use it three ways:

- Double-click `SubMiner mpv` to open `mpv` with the SubMiner profile.
- Drag a video file onto `SubMiner mpv` to launch that file with the same profile.
- Run it directly from Command Prompt or PowerShell with `--launch-mpv`.

```powershell
& "C:\Program Files\SubMiner\SubMiner.exe" --launch-mpv
& "C:\Program Files\SubMiner\SubMiner.exe" --launch-mpv "C:\Videos\episode 01.mkv"
```

This flow requires `mpv.exe` to be on `PATH`. If it is installed elsewhere, set `SUBMINER_MPV_PATH` to the full `mpv.exe` path before launching.

### Launcher Subcommands

- `subminer jellyfin` / `subminer jf`: Jellyfin-focused workflow aliases.
- `subminer yt` / `subminer youtube`: YouTube-focused shorthand flags (`-o`, `--keep-temp`, `--whisper-*`).
- `subminer doctor`: health checks for core dependencies and runtime paths.
- `subminer config`: config helpers (`path`, `show`).
- `subminer mpv`: mpv helpers (`status`, `socket`, `idle`).
- `subminer dictionary <path>`: generates a Yomitan-importable character dictionary ZIP from a file/directory target.
- `subminer texthooker`: texthooker-only shortcut (same behavior as `--texthooker`).
- `subminer app` / `subminer bin`: direct passthrough to the SubMiner binary/AppImage.
- Subcommand help pages are available (for example `subminer jellyfin -h`, `subminer yt -h`).

### First-Run Setup

SubMiner auto-opens the setup popup on fresh installs when launched with `--start` or `--background` and setup is incomplete.

You can also open it manually:

```bash
subminer --setup
SubMiner.AppImage --setup
```

Setup flow:

- config file: create the default config directory and prefer `config.jsonc`
- plugin status: install or skip the bundled mpv plugin
- Yomitan shortcut: open bundled Yomitan settings directly from the setup window
- dictionary check: ensure at least one bundled Yomitan dictionary is available
- Windows: optionally create or remove `SubMiner mpv` Start Menu/Desktop shortcuts (`SubMiner.exe --launch-mpv`)
- refresh: re-check plugin + dictionary state without restarting
- `Finish setup` stays disabled until dictionary availability is detected
- finish action writes setup completion state and suppresses future auto-open prompts

AniList character dictionary auto-sync (optional):

- Enable with `anilist.characterDictionary.enabled=true` in config.
- SubMiner syncs the currently watched AniList media into a per-media snapshot, then rebuilds one merged `SubMiner Character Dictionary` from the most recently used snapshots.
- Rotation limit defaults to 3 recent media snapshots in that merged dictionary (`maxLoaded`).

Use subcommands for Jellyfin/YouTube command families (`subminer jellyfin ...`, `subminer yt ...`).
Top-level launcher flags like `--jellyfin-*` and `--yt-subgen-*` are intentionally rejected.

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

You can define a matching profile in `~/.config/mpv/mpv.conf` for consistency when launching `mpv` manually or from other tools. `subminer` launches with `--profile=subminer` by default (or override with `subminer -p <profile> ...`):

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

::: warning
`secondary-slang` is not a valid mpv option. Use `slang` with `sid=auto` / `secondary-sid=auto` to set subtitle language preferences.
:::

### Yomitan setup

SubMiner includes a bundled Yomitan extension for overlay word lookup. This bundled extension is separate from any Yomitan browser extension you may have installed.

For SubMiner overlay lookups to work, open Yomitan settings (`subminer app --settings` or `SubMiner.AppImage --settings`) and import at least one dictionary in the bundled Yomitan instance.

If you also use Yomitan in a browser, configure that browser profile separately; it does not inherit dictionaries or settings from the bundled instance.

### YouTube Playback

`subminer` accepts direct URLs (for example, YouTube links) and `ytsearch:` targets.
For YouTube playback, SubMiner now generates or downloads subtitle tracks before mpv starts, then launches mpv with the resolved subtitle files attached.

Notes:

- Install `yt-dlp` so mpv can resolve YouTube streams and subtitle tracks reliably.
- For YouTube URLs, `subminer` now generates any missing subtitles before mpv launch.
- It probes manual/native YouTube subtitle tracks first, then falls back to local `whisper.cpp` only for missing tracks.
- Primary subtitle target languages come from `youtubeSubgen.primarySubLanguages` (defaults to `["ja","jpn"]`).
- Secondary target languages come from `secondarySub.secondarySubLanguages` (defaults to English if unset).
- Whisper translation fallback currently only supports English secondary targets; non-English secondary targets rely on native/manual subtitle availability.
- Optional AI cleanup for whisper-generated subtitles is controlled by `youtubeSubgen.fixWithAi` plus the shared top-level `ai` config (with optional `youtubeSubgen.ai` overrides).
- Configure defaults in `$XDG_CONFIG_HOME/SubMiner/config.jsonc` (or `~/.config/SubMiner/config.jsonc`) under `youtubeSubgen`, `secondarySub`, and `ai`.
- CLI overrides are available through `subminer yt` / `subminer youtube`:
  - `-o, --out-dir`
  - `--keep-temp`
  - `--whisper-bin`
  - `--whisper-model`
  - `--whisper-vad-model`
  - `--whisper-threads`
  - `--yt-subgen-audio-format`

## Keybindings

See [Keyboard Shortcuts](/shortcuts) for the full reference, including mining shortcuts, overlay controls, and customization.

**Global shortcuts** (work system-wide):

| Keybind       | Action                 |
| ------------- | ---------------------- |
| `Alt+Shift+O` | Toggle visible overlay |
| `Alt+Shift+Y` | Open Yomitan settings  |

::: tip
`Alt+Shift+Y` is fixed and not configurable. All other shortcuts can be changed under `shortcuts` in your config.
:::

Hovering over subtitle text pauses mpv by default; leaving resumes it. Disable with `subtitleStyle.autoPauseVideoOnHover: false`. To also pause while the Yomitan popup is open, set `subtitleStyle.autoPauseVideoOnYomitanPopup: true`.

### Drag-and-Drop

- Drop video files onto the overlay to replace current playback.
- Hold `Shift` while dropping to append to the playlist instead.

Next: [Mining Workflow](/mining-workflow) — word lookup, card creation, and the full mining loop.
