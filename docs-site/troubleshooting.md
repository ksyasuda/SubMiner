# Troubleshooting

Find your symptom below. If you saw an error message, search this page for its text.

## Diagnose first

Run the launcher's dependency check:

```bash
subminer doctor
```

It reports the app binary, `mpv`, `yt-dlp`, `ffmpeg`, `fzf`, `rofi`, your config file, and the mpv socket path. It exits non-zero when the app binary or `mpv` is missing.

Logs are written to daily files:

| Platform      | Log directory              |
| ------------- | -------------------------- |
| Linux / macOS | `~/.config/SubMiner/logs/` |
| Windows       | `%APPDATA%\SubMiner\logs\` |

Files are named `app-<date>.log`, `launcher-<date>.log`, and `mpv-<date>.log`. The mpv log is off by default. Turn log files on or off and set retention under [`logging`](/configuration#logging).

For more detail, raise the log level for one run:

```bash
subminer --log-level debug video.mkv
SubMiner.AppImage --start --log-level debug
```

The default level is `warn`. `--dev` and `--debug` switch the app into dev mode but do not change log verbosity. To inspect the overlay itself, focus it and press `y` then `d` to open DevTools.

## Overlay starts but shows no subtitles

SubMiner reads subtitles from mpv over an IPC socket (a named pipe on Windows). If the paths do not match, the overlay appears but stays empty.

- The `subminer` launcher sets the socket for you. If you start mpv yourself, pass `--input-ipc-server=/tmp/subminer-socket`.
- If you changed `mpv.socketPath`, use the same path in your mpv config.

SubMiner reconnects on its own if mpv restarts.

## Overlay does not appear

- Confirm SubMiner is running (`SubMiner.AppImage --start`, or check for the process).
- Linux: Hyprland and Sway work natively. Any other compositor needs mpv and SubMiner under X11 or Xwayland, with `xdotool`, `xprop`, and `xwininfo` installed. See [KDE Plasma and other Wayland compositors](#kde-plasma-and-other-wayland-compositors).
- macOS: grant Accessibility permission in System Settings > Privacy & Security > Accessibility.

## Overlay is on the wrong monitor or position

SubMiner follows the mpv window. Tracking needs `hyprctl` (Hyprland), `swaymsg` (Sway), or `xdotool` and `xwininfo` (X11) on `PATH`.

If the position is only slightly off, right-click and drag the subtitle text to adjust the offset.

## Clicks pass through the overlay

- The overlay only takes input while the cursor is over subtitle text. Hover the text directly.
- Toggle the overlay off and on with `Alt+Shift+O`.
- Linux: if clicks keep failing, toggle the overlay off, click the mpv window, then toggle it back on.

## Hovering a word shows no popup

If you have not set up dictionaries yet, start with [Yomitan setup](/usage#yomitan-setup).

- Open Yomitan settings (`Alt+Shift+Y` or `SubMiner.AppImage --yomitan`) and confirm at least one dictionary is imported and enabled.
- If `yomitan.externalProfilePath` is set, manage dictionaries in that external profile. SubMiner opens it read-only and has no settings window of its own in that mode.
- Check the log for "Loaded Yomitan extension".

Word boundaries come from Yomitan's parser. Some splits will be wrong, since Japanese has no spaces.

## "Yomitan extension not found in any search path"

The bundled Yomitan is missing. Re-download the AppImage, or place an unpacked Yomitan extension in `~/.config/SubMiner/yomitan`. Source builds must run `bun run build` first to produce `build/yomitan`.

## "MeCab not found on system"

This is informational. Tokenization uses Yomitan, not MeCab. Install MeCab only if you want to silence the message:

- Arch: `sudo pacman -S mecab mecab-ipadic`
- Ubuntu/Debian: `sudo apt install mecab libmecab-dev mecab-ipadic-utf8`
- macOS: `brew install mecab mecab-ipadic`

## "AnkiConnect: unable to connect"

Anki must be running with the AnkiConnect add-on. See [Anki integration prerequisites](/anki-integration#prerequisites).

- Direct mode: check that `ankiConnect.url` matches the AnkiConnect port.
- Proxy mode: check `ankiConnect.proxy.upstreamUrl`, and point external Yomitan or browser clients at the SubMiner proxy.

SubMiner keeps retrying and logs "AnkiConnect connection restored" once Anki is back.

## Cards are created but fields are empty

Each name in `ankiConnect.fields` must match a field on your note type. Matching is case-insensitive, but otherwise the spelling must match. Unknown fields are skipped without an error. For example, config `Audio` does not fill a note field named `SentenceAudio`. See [Anki integration](/anki-integration).

## "Update failed" when mining

The card was deleted in Anki before SubMiner finished enriching it, or the note type changed and a mapped field no longer exists.

## "Subtitle timing not found; copy again while playing"

SubMiner has no timing for the current line yet. This happens when paused before any subtitle arrived, after switching subtitle tracks, or while an external subtitle file is still loading. Resume playback, wait for the next line, and mine again.

## "FFmpeg not found"

Audio clips and screenshots need FFmpeg. Without it, cards are still created with empty media fields.

- Arch: `sudo pacman -S ffmpeg`
- Ubuntu/Debian: `sudo apt install ffmpeg`
- macOS: `brew install ffmpeg`

## Audio or screenshot generation is slow or times out

- Use a local copy if the video is on a slow network mount.
- Set `ankiConnect.media.imageType` to `"static"`. Animated AVIF is the slowest path.
- Lower `ankiConnect.media.imageQuality` or `ankiConnect.media.maxMediaDuration`.

## Subtitle sync (subsync)

Subtitle sync needs at least one of alass or ffsubsync. Neither ships with SubMiner.

**"Configured alass executable not found"**: install it (`paru -S alass` or `cargo install alass-cli`), or set `subsync.alass_path`.

**"Configured ffsubsync executable not found"**: install it (`paru -S python-ffsubsync` or `pip install ffsubsync`), or set `subsync.ffsubsync_path`.

**"alass synchronization failed" / "ffsubsync synchronization failed"**:

- alass needs a reference: a second subtitle track or the local video file. It cannot use the track being retimed.
- `ffmpeg` must be installed to extract internal subtitle tracks.
- ffsubsync only works on local files, not streams.
- Run the tool by hand to see its full error output.

## "xz binary not found"

TsukiHime subtitles are xz-compressed. Install `xz`:

- Arch: `sudo pacman -S xz`
- Ubuntu/Debian: `sudo apt install xz-utils`
- Fedora: `sudo dnf install xz`
- macOS: `brew install xz`
- Windows: `scoop install main/xz`, or download XZ Utils from [tukaani.org/xz](https://tukaani.org/xz/) and add the folder with `xz.exe` to `PATH`. Restart SubMiner afterwards.

Other TsukiHime errors are covered in [TsukiHime integration](/tsukihime-integration#troubleshooting).

## Jimaku

**"Jimaku request failed" or HTTP 429**: you hit the Jimaku rate limit. Wait for the time shown in the message. Setting `jimaku.apiKey` or `jimaku.apiKeyCommand` gives you a higher limit.

## Character names are not highlighted

See [Character dictionary](/character-dictionary) for the full list. The common causes:

- `subtitleStyle.nameMatchEnabled` is off, or the media did not resolve to an AniList entry.
- Portraits need `subtitleStyle.nameMatchImagesEnabled`.
- Wrong characters: open the manager (`Ctrl/Cmd+D`) and use **Override** to pick the right AniList entry.
- The feature is disabled when `yomitan.externalProfilePath` is set.

## "Failed to register global shortcut"

Another app or your desktop already uses `Alt+Shift+Y` (Yomitan settings). Free it in your desktop or window manager settings. On Hyprland, add a `pass` rule (see [Hyprland](#hyprland)).

## Overlay shortcuts do nothing

Overlay shortcuts only work while the overlay has focus. Click the overlay, or press `Alt+Shift+O` with mpv or the overlay focused.

## Update checks

**"Update check failed"**: GitHub could not be reached. Check your connection and retry from the tray menu or with `subminer -u`.

**No prerelease offered**: set `updates.channel` to `"prerelease"` to include beta and RC builds.

**Launcher update shows a sudo command**: the launcher lives in a protected path such as `/usr/local/bin`. Run the command shown to replace it.

## Playback feels sluggish

Idle playback is cheap. Load comes from the first tokenization burst, media generation, subtitle sync, and Anki enrichment. To cut it:

- `ankiConnect.media.imageType: "static"`, plus lower `imageQuality` and `maxMediaDuration`.
- `subtitleStyle.enableJlpt: false` and `subtitleStyle.frequencyDictionary.enabled: false`.
- `secondarySub.defaultMode: "hidden"`.
- `immersionTracking.enabled: false` to stop stats logging.

Also check that only one SubMiner instance is running, and whether `ffmpeg`, `yt-dlp`, or a sync tool is the process using CPU.

## Linux

### Tray icon missing

Linux trays need a StatusNotifier/AppIndicator host. Hyprland has none by default. Enable a tray in Waybar, Hyprpanel, or another panel.

### Hyprland

SubMiner applies float, no-border, and no-blur properties to its window itself. If the overlay still has a border or an opaque background, a global opacity or blur rule is usually overriding it. Add a rule for the `SubMiner` class. Lua config:

```lua
hl.window_rule({
  match = { class = "^SubMiner$" },
  float = true,
  border_size = 0,
  xray = false,
  no_shadow = true,
  no_blur = true,
  no_dim = true,
  opaque = true,
  dim_around = false,
  opacity = "1.0 override 1.0 override",
})
```

Older `hyprland.conf` configs:

```ini
windowrule = float on, match:class SubMiner
windowrule = border_size 0, match:class SubMiner
windowrule = xray off override, match:class SubMiner
windowrule = no_shadow on, match:class SubMiner
windowrule = no_blur on, match:class SubMiner
```

Hyprland swallows global shortcuts unless you pass them through. Add a `pass` bind for each one, and update it if you remap the key:

```ini
bind = ALT SHIFT, O, pass, class:^(SubMiner)$
bind = ALT SHIFT, Y, pass, class:^(SubMiner)$
```

If the overlay stays behind fullscreen mpv, check that the mpv socket is connected and that `hyprctl -j clients` works from the environment that launched SubMiner.

See the Hyprland wiki on [global keybinds](https://wiki.hypr.land/Configuring/Binds/#global-keybinds) and [window rules](https://wiki.hypr.land/Configuring/Window-Rules/).

### KDE Plasma and other Wayland compositors

Outside Hyprland and Sway, Wayland does not let the overlay stay on top of mpv, so both must run under Xwayland. SubMiner does this automatically for itself and for every mpv it launches (launcher, tray, Jellyfin, YouTube). Install `xdotool`, `xprop`, and `xwininfo`.

**Overlay sits behind mpv, and hover or Yomitan stops working**: mpv started as a native Wayland window. This happens when you launch mpv yourself. Launch through SubMiner, or force Xwayland in your own command:

```bash
mpv --gpu-context=x11vk,x11egl,x11 video.mkv
```

`WAYLAND_DISPLAY= mpv video.mkv` also works, as does `gpu-context=x11vk` (or `x11egl`) in `mpv.conf`. To check, `xdotool search --class mpv` should print a window id.

**Overlay stays above an unrelated app**: SubMiner can only see X11/Xwayland windows in this mode. Run that app under Xwayland too.

## macOS

- Accessibility permission is required for window tracking: System Settings > Privacy & Security > Accessibility.
- If Gatekeeper blocks the app, right-click it and choose Open, or run `xattr -d com.apple.quarantine /path/to/SubMiner.app`.
