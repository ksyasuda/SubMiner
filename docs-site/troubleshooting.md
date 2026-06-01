# Troubleshooting

Common issues and how to resolve them. Most problems fall into one of a few buckets — the overlay shows but subtitles don't (see [MPV Connection](#mpv-connection)), cards aren't being created or come out empty (see [AnkiConnect](#ankiconnect)), or word lookups don't appear (see [Yomitan](#yomitan)). If an error message popped up on screen, search this page for the exact text — most headings below are quoted error strings.

## MPV Connection

**Overlay starts but shows no subtitles**

SubMiner connects to mpv via a Unix socket (or named pipe on Windows). If the socket does not exist or the path does not match, the overlay will appear but subtitles will never arrive.

- Ensure mpv is running with `--input-ipc-server=/tmp/subminer-socket`.
- If you use a custom socket path, set it in both your mpv config and SubMiner config (`mpv.socketPath`).
- The `subminer` wrapper script sets the socket automatically when it launches mpv. If you launch mpv yourself, the `--input-ipc-server` flag is required.

SubMiner retries the connection automatically with increasing delays (200 ms, 500 ms, 1 s, 2 s on first connect; 1 s, 2 s, 5 s, 10 s on reconnect). If mpv exits and restarts, the overlay reconnects without needing a restart.

If the overlay never appears at all, see [Playback Startup Flow](./architecture#playback-startup-flow) for how a managed launch starts mpv and brings up the overlay.

## Logging and App Mode

- Default log output is `warn`.
- Use `--log-level` for more/less output.
- Use `--dev`/`--debug` only to force app/dev mode (for example to get dev behavior from the overlay/app); they do not change log verbosity.
- You can combine both, for example `SubMiner.AppImage --start --dev --log-level debug`, when you need maximum diagnostics.

## Performance and Resource Impact

### At a glance

- Baseline: `SubMiner --start` is usually lightweight for normal playback.
- Common spikes come from:
  - first subtitle parse/tokenization bursts
  - media generation (`ffmpeg` audio/image and AVIF paths)
  - media sync and subtitle tooling (`alass`, `ffsubsync`)
  - `ankiConnect` enrichment (plus polling overhead when proxy mode is disabled)

### If playback feels sluggish

1. Reduce overlay workload:

- set secondary subtitles hidden:
  - `secondarySub.defaultMode: "hidden"`
- disable optional enrichment:
  - `subtitleStyle.enableJlpt: false`
  - `subtitleStyle.frequencyDictionary.enabled: false`

2. Reduce rendering pressure:

- lower `subtitleStyle.css["font-size"]`
- keep overlay complexity minimal during heavy CPU periods

3. Reduce media overhead:

- keep `ankiConnect.media.imageType` set to `static` (avoid animated AVIF unless needed)
- lower `ankiConnect.media.imageQuality`
- reduce `ankiConnect.media.maxMediaDuration`

4. Lower integration cost:

- disable AI translation when not needed (`ankiConnect.ai.enabled: false`)
- if needed, run immersion telemetry with lower duration expectations (`immersionTracking.enabled: false` for constrained sessions)
- favor the default lightweight YouTube subtitle startup settings on low-resource systems

### Practical low-impact profile

```json
{
  "subtitleStyle": {
    "css": {
      "font-size": "30px"
    },
    "enableJlpt": false,
    "frequencyDictionary": {
      "enabled": false
    }
  },
  "secondarySub": {
    "defaultMode": "hidden"
  },
  "ankiConnect": {
    "media": {
      "imageType": "static",
      "imageQuality": 80,
      "maxMediaDuration": 12
    },
    "ai": {
      "enabled": false
    }
  },
  "immersionTracking": {
    "enabled": false
  }
}
```

### If usage is still high

- Confirm only one SubMiner instance is running.
- Check whether bottlenecks are `ffmpeg`, `yt-dlp`, or sync tooling in system monitor.
- Keep the default `warn` level for normal use; raise to `info` or `debug` only for targeted diagnosis.
- Reproduce once with `SubMiner.AppImage --start --log-level debug` and open DevTools (`y` then `d`) if freezes recur.

**"Failed to parse MPV message"**

Logged when a malformed JSON line arrives from the mpv socket. Usually harmless — SubMiner skips the bad line and continues. If it happens constantly, check that nothing else is writing to the same socket path.

## Updates

**"Update check failed"**

Manual update checks show this when GitHub Releases or updater metadata cannot be reached. Check your network connection, then try again from the tray menu or:

```bash
subminer -u
```

Automatic checks log failures quietly so playback is not interrupted.

**"SubMiner is up to date" but a prerelease exists**

SubMiner uses the configured release channel for update checks. Set `updates.channel` to `"prerelease"` in `config.jsonc` when you want update checks to include beta and RC releases.

**Launcher update shows a sudo command**

The detected launcher is installed in a protected path such as `/usr/local/bin/subminer` or `/usr/bin/subminer`. SubMiner does not elevate itself. Run the command shown in the popup to replace the launcher after checksum verification.

**OSD update notification did not appear**

`updates.notificationType: "osd"` uses the existing mpv/overlay notification path. If mpv is disconnected, SubMiner logs the update and does not force-start the overlay. Use `"system"` or `"both"` if you want OS notifications outside playback.

## AnkiConnect

**"AnkiConnect: unable to connect"**

First confirm you've completed the [Anki Integration prerequisites](/anki-integration#prerequisites) — Anki must be running with the AnkiConnect add-on installed.

SubMiner connects to the active Anki endpoint:

- `ankiConnect.url` (direct mode, default `http://127.0.0.1:8765`)
- `http://<ankiConnect.proxy.host>:<ankiConnect.proxy.port>` (proxy mode)

This error means the active endpoint is unavailable, or (in proxy mode) the proxy cannot reach `ankiConnect.proxy.upstreamUrl`.

- If you changed the AnkiConnect port, update `ankiConnect.url` (or `ankiConnect.proxy.upstreamUrl` if using proxy mode).
- If using external Yomitan/browser clients, confirm they point to your SubMiner proxy URL.

SubMiner retries with exponential backoff (up to 5 s) and suppresses repeated error logs after 5 consecutive failures. When Anki comes back, you will see "AnkiConnect connection restored".

**Cards are created but fields are empty**

Field names in your config must match your Anki note type exactly (case-sensitive). Check `ankiConnect.fields` — for example, if your note type uses `SentenceAudio` but your config says `Audio`, the field will not be populated.

See [Anki Integration](/anki-integration) for the full field mapping reference.

**"Update failed" OSD message**

Shown when SubMiner tries to update a card that no longer exists, or when AnkiConnect rejects the update. Common causes:

- The card was deleted in Anki between creation and enrichment update.
- The note type changed and a mapped field no longer exists.

## Overlay

**Overlay does not appear**

- Confirm SubMiner is running: `SubMiner.AppImage --start` or check for the process.
- On Linux, the overlay requires a supported window backend. Hyprland and Sway have native Wayland support; all other compositors require both mpv and SubMiner to run under X11 or Xwayland (`xdotool` and `xwininfo` must be installed).
- On macOS, grant Accessibility permission to SubMiner in System Settings > Privacy & Security > Accessibility.

**Overlay appears but clicks pass through / cannot interact**

- Make sure you are hovering over subtitle text — the overlay only becomes interactive when the cursor is over a subtitle.
- On macOS/Windows: toggle the overlay off and back on (`Alt+Shift+O`) to re-enable pointer events.
- On Linux: mouse event handling is unreliable in some Electron/compositor combinations. If clicks consistently fail, toggle the overlay off, click the underlying mpv window, then toggle it back on.

**Overlay briefly freezes after a modal/runtime error**

- Renderer errors now trigger an automatic recovery path. You should see a short toast ("Renderer error recovered. Overlay is still running.").
- Recovery closes any open modal and restores click-through/shortcuts automatically without interrupting mpv playback.
- If errors keep recurring, toggle the overlay's DevTools using overlay chord `y` then `d` (or global `F12`) and inspect the `renderer overlay recovery` error payload for stack trace + modal/subtitle context.

**Overlay is on the wrong monitor or position**

SubMiner positions the overlay by tracking the mpv window. If tracking fails:

- Hyprland: Ensure `hyprctl` is available.
- Sway: Ensure `swaymsg` is available.
- X11: Ensure `xdotool` and `xwininfo` are installed.

If the overlay position is slightly off, right-click and drag on subtitle text to fine-tune the overlay subtitle offset.

## Yomitan

If you haven't set up dictionaries yet, see [Yomitan setup](/usage#yomitan-setup) first.

**"Yomitan extension not found in any search path"**

SubMiner bundles Yomitan and searches for it in these locations (in order):

1. `build/yomitan` (local/source build output)
2. `<resources>/yomitan` (Electron resources path)
3. `/usr/share/SubMiner/yomitan`
4. `~/.config/SubMiner/yomitan` (user-data fallback on Linux)

SubMiner does not load the source tree directly from `vendor/subminer-yomitan`; source builds must produce `build/yomitan` first.

If you installed from the AppImage and see this error, the package may be incomplete. Re-download the AppImage or place the unpacked Yomitan extension manually in `~/.config/SubMiner/yomitan`.

**Yomitan lookup popup does not appear when hovering words or triggering lookup**

- Verify Yomitan loaded successfully — check the terminal output for "Loaded Yomitan extension".
- Yomitan requires dictionaries to be installed. Open Yomitan settings (`Alt+Shift+Y` or `SubMiner.AppImage --yomitan`) and confirm at least one dictionary is imported.
- If `yomitan.externalProfilePath` is set, import/check dictionaries in the external app/profile instead. SubMiner treats that profile as read-only and does not open its own Yomitan settings window.
- If the overlay shows subtitles but hover lookup never resolves on tokens, the tokenizer may have failed. See the MeCab section below.

## MeCab / Tokenization

**"MeCab not found on system"**

This is informational, not an error. SubMiner tokenization is driven by Yomitan's internal parser. MeCab availability checks may still run for auxiliary token metadata, but MeCab is not used as a tokenization fallback path.

To install MeCab:

- **Arch Linux**: `sudo pacman -S mecab mecab-ipadic`
- **Ubuntu/Debian**: `sudo apt install mecab libmecab-dev mecab-ipadic-utf8`
- **macOS**: `brew install mecab mecab-ipadic`

**Words are not segmented correctly**

Japanese word boundaries depend on Yomitan parser output. If segmentation seems wrong:

- Verify Yomitan dictionaries are installed and active.
- Note that CJK characters without spaces are segmented using parser heuristics, which is not always perfect.

## Character Dictionary

Character names from AniList are matched and highlighted in subtitles via the bundled Yomitan. See [Character Dictionary](/character-dictionary) for setup and the full troubleshooting list — the most common issues:

- **Names not highlighting:** Confirm `subtitleStyle.nameMatchEnabled` is `true`, and that the current media resolved to an AniList entry (SubMiner needs a media ID to fetch characters). No AniList account or token is required — character data uses public GraphQL queries.
- **Inline portraits missing:** Confirm `subtitleStyle.nameMatchImagesEnabled` is `true`. Portraits also require AniList to return an image and the download to succeed during snapshot generation.
- **Wrong characters showing:** Open the in-app manager (`Ctrl/Cmd+D`) and use **Override** to pin the correct AniList match for the series.
- **Feature unavailable:** If `yomitan.externalProfilePath` is set, SubMiner runs in read-only external-profile mode and its character-dictionary features are disabled.

## Media Generation

**"FFmpeg not found"**

SubMiner uses FFmpeg to extract audio clips and generate screenshots. Install it:

- **Arch Linux**: `sudo pacman -S ffmpeg`
- **Ubuntu/Debian**: `sudo apt install ffmpeg`
- **macOS**: `brew install ffmpeg`

Without FFmpeg, card creation still works but audio and image fields will be empty.

**Audio or screenshot generation hangs**

Media generation has a 30-second timeout (60 seconds for animated AVIF). If your video file is on a slow network mount or the codec requires software decoding, generation may time out. Try:

- Using a local copy of the video file.
- Reducing `media.imageQuality` or switching from `avif` to `static` image type.
- Checking that `media.maxMediaDuration` is not set too high.

## Shortcuts

**"Failed to register global shortcut"**

Global shortcuts (`Alt+Shift+O`, `Alt+Shift+Y`) may conflict with other applications or desktop environment keybindings.

- Check your DE/WM keybinding settings for conflicts.
- Change the shortcut in your config under `shortcuts.toggleVisibleOverlayGlobal`.
- On Wayland, global shortcut registration has limitations depending on the compositor. Only Hyprland and Sway are supported natively — see the [Hyprland](#hyprland) section below for shortcut passthrough rules. Other Wayland compositors require X11/Xwayland.

**Overlay keybindings not working**

Overlay-local shortcuts (Space, arrow keys, etc.) only work when the overlay window has focus. Click on the overlay or use the global shortcut to toggle it to give it focus.

## Subtitle Timing

**"Subtitle timing not found; copy again while playing"**

This OSD message appears when you try to mine a sentence but SubMiner has no timing data for the current subtitle. Causes:

- The video is paused and no subtitle has been received yet.
- The subtitle track changed and timing data was cleared.
- You are using an external subtitle file that mpv has not fully loaded.

Resume playback and wait for the next subtitle to appear, then try mining again.

## Subtitle Sync (Subsync)

Both **alass** and **ffsubsync** are optional external dependencies. Subtitle syncing requires at least one of them to be installed.

**"Configured alass executable not found"**

Install alass or configure the path:

- **Arch Linux (AUR)**: `paru -S alass`
- **Cargo**: `cargo install alass-cli`
- Set the path: `subsync.alass_path` in your config.

**"Configured ffsubsync executable not found"**

Install ffsubsync or configure the path:

- **Arch Linux (AUR)**: `paru -S python-ffsubsync`
- **pip**: `pip install ffsubsync`
- Must be on `PATH` or configured via `subsync.ffsubsync_path` in your config.

**"Subtitle synchronization failed"**

If subtitle sync fails:

- Ensure the reference subtitle track exists in the video (alass requires a source track).
- Check that `ffmpeg` is available (used to extract the internal subtitle track).
- Try running the sync tool manually to see detailed error output.
- ffsubsync requires local files and cannot handle remote media streams (e.g., streaming URLs).

## Jimaku

**"Jimaku request failed" or HTTP 429**

The Jimaku API has rate limits. If you see 429 errors, wait for the retry duration shown in the OSD message and try again. If you have a Jimaku API key, set it in `jimaku.apiKey` or `jimaku.apiKeyCommand` to get higher rate limits.

## Platform-Specific

### Linux

- **Wayland (Hyprland/Sway only)**: Native Wayland support is limited to Hyprland and Sway. Window tracking uses compositor-specific commands (`hyprctl` / `swaymsg`). If these are not on `PATH`, tracking will fail silently. Other Wayland compositors (KDE Plasma, GNOME, …) are not supported natively — both mpv and SubMiner must run under X11 or Xwayland instead. On those sessions SubMiner forces XWayland automatically for itself and for every mpv it launches (see [KDE Plasma & other Wayland compositors](#kde-plasma--other-wayland-compositors)).
- **X11 / Xwayland**: Requires `xdotool`, `xprop`, and `xwininfo`. If missing, the overlay cannot track the mpv window position. This is the required backend for any Wayland compositor other than Hyprland or Sway — both mpv and SubMiner must be running under X11/Xwayland for window tracking _and_ for the overlay to stay above mpv (Wayland forbids clients from controlling window stacking). SubMiner uses a managed X11 overlay while mpv is windowed, switches to an override-redirect X11 overlay while tracked mpv is fullscreen, and hides/releases that overlay when another X11/Xwayland app takes focus. The visible overlay stays hidden until SubMiner has tracked mpv geometry, so startup should not create a display-sized fallback overlay while tokenization warms up.
- **Tray icon missing**: SubMiner creates an Electron tray icon in `--background` mode, but Linux trays require a StatusNotifier/AppIndicator host. Hyprland does not provide one by itself; enable a tray in Waybar, Hyprpanel, or another panel. If Electron cannot register the tray, SubMiner logs a warning that mentions the missing tray host.
- **Mouse passthrough**: On Linux X11/Xwayland, SubMiner uses `xdotool` to poll the cursor and only enables overlay input while the cursor is over subtitle or popup regions. Outside those regions, pointer input passes through to mpv. Native Wayland compositors other than Hyprland/Sway cannot provide the stacking control SubMiner needs.

### Hyprland

SubMiner's overlay is a transparent, frameless Electron window that must be kept above mpv. SubMiner tries to apply the floating, borderless, no-shadow, and no-blur properties itself each time it places the overlay. It detects Hyprland's active config provider and uses Lua `hl.dsp.window.*` dispatchers for recent Hyprland Lua configs, or the legacy dispatcher syntax for older hyprlang configs. On many configurations that is enough, but if your Hyprland version doesn't honor those runtime dispatches — or a broad rule in your config forces opacity/blur on every window — add explicit window rules so the overlay is exempt. You also need `pass` bindings to forward global shortcuts to SubMiner (see below).

**Overlay is not transparent or has a visible border**

Add a window rule matching SubMiner's window class. Recent Hyprland uses the Lua config format:

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

On older Hyprland releases that still use the hyprlang config (`hyprland.conf`), use the equivalent `windowrule` lines:

```ini
windowrule = float on, match:class SubMiner
windowrule = border_size 0, match:class SubMiner
windowrule = xray off override, match:class SubMiner
windowrule = no_shadow on, match:class SubMiner
windowrule = no_blur on, match:class SubMiner
```

If you still see a solid background or visual artifacts instead of the mpv video underneath, the culprit is almost always a global opacity/blur rule applying to the overlay — the `opaque`/`opacity` and `no_blur` fields above override it.

**Global shortcuts not working**

On Hyprland, Electron cannot register global shortcuts on its own. You must explicitly pass keybindings to SubMiner using `pass` rules:

```ini
bind = ALT SHIFT, O, pass, class:^(SubMiner)$
bind = ALT SHIFT, Y, pass, class:^(SubMiner)$
```

Add a `pass` rule for each global shortcut you configure. The defaults are `Alt+Shift+O` (toggle overlay) and `Alt+Shift+Y` (Yomitan settings). If you remap `shortcuts.toggleVisibleOverlayGlobal` to a different key, update the `pass` rule to match.

Without these rules, Hyprland intercepts the keypresses before they reach SubMiner, and the shortcuts silently do nothing.

**Overlay stays behind mpv after fullscreen**

SubMiner watches mpv's `fullscreen` property and refreshes the overlay geometry when it changes. If the overlay still does not move or rise above fullscreen mpv, confirm that the mpv IPC socket is connected and that `hyprctl -j clients` and `hyprctl -j monitors` work from the same environment that launched SubMiner.

For more details, see the Hyprland docs on [global keybinds](https://wiki.hypr.land/Configuring/Binds/#global-keybinds) and [window rules](https://wiki.hypr.land/Configuring/Window-Rules/).

### KDE Plasma & other Wayland compositors

On any Wayland session that is not Hyprland or Sway (KDE Plasma, GNOME, and others), the overlay can only stay above mpv when both processes run under **XWayland** — the Wayland protocol forbids clients from controlling window stacking, so the overlay's "always on top" becomes a no-op on a native Wayland surface.

SubMiner handles this automatically:

- It launches its own window under XWayland (it sets `--ozone-platform-hint=x11`).
- Every mpv it launches (via the `subminer` launcher, Jellyfin, or YouTube) is pinned to XWayland too — Wayland environment hints are stripped and an X11 GPU context (`--gpu-context=x11egl,x11`) is applied.
- While mpv is windowed, the overlay is a managed X11 window owned by the tracked mpv window (`WM_TRANSIENT_FOR`), so it stays above mpv while other foreground X11/Xwayland apps can still cover both windows.
- While tracked mpv is fullscreen, SubMiner swaps the visible overlay to a focusable-false X11 override-redirect window. That path can stay above the active fullscreen mpv window without requiring a KDE/KWin-specific rule, and SubMiner hides/releases it when mpv is no longer the active X11/Xwayland window.
- The visible overlay is shown inactive on Linux, so normal hover should not steal keyboard focus from mpv.
- During startup and fullscreen transitions, SubMiner waits for tracked mpv geometry before showing the visible overlay and skips the fullscreen restack hide/show path after mpv leaves fullscreen. That avoids a temporary full-screen overlay or black window while the subtitle tokenizer and Yomitan warmups finish.
- If the subtitle sidebar is open during a windowed/fullscreen transition, SubMiner restores it on the replacement overlay window. Subtitle hit regions are also refreshed as soon as the first measured subtitle line is reported, so hover and Yomitan lookup should work on the first visible line.

Requirements: `xdotool`, `xprop`, and `xwininfo` must be installed. SubMiner uses root `_NET_ACTIVE_WINDOW` from `xprop` for focus detection and falls back to `xdotool getactivewindow` when that signal is unavailable.

**Overlay sits behind mpv / pause-on-hover and Yomitan stop working**

This almost always means mpv came up as a **native Wayland** window that the XWayland overlay cannot cover. It happens when mpv is launched **manually** (your own command), because SubMiner can only force XWayland on the mpv processes it launches itself. Fix it one of these ways:

- Launch playback through SubMiner (the `subminer` launcher or the tray), which forces XWayland for you, or
- Force XWayland in your own mpv invocation, e.g. `mpv --gpu-context=x11egl …`, or launch with `WAYLAND_DISPLAY= mpv …`, or set `gpu-context=x11egl` in your `mpv.conf`.

To confirm mpv is on XWayland, `xdotool search --class mpv` should return a window id (a native Wayland mpv returns nothing).

**Overlay stays above an unrelated foreground app**

SubMiner can only detect focus for X11/Xwayland windows in this mode. If a native Wayland app covers mpv but the overlay stays visible, run that app under Xwayland too or use Hyprland/Sway native support. Generic X11 cannot observe native Wayland foreground windows.

### macOS

- **Accessibility permission**: Required for window tracking. Grant it in System Settings > Privacy & Security > Accessibility.
- **Gatekeeper**: If macOS blocks SubMiner, right-click the app and select "Open" to bypass the warning, or remove the quarantine attribute: `xattr -d com.apple.quarantine /path/to/SubMiner.app`

## See Also

Feature-specific issues are covered in each feature's own page:

- [Anki Integration](/anki-integration) — card creation, field mapping, and AnkiConnect setup
- [AniList Integration](/anilist-integration) — watch-progress sync and authentication
- [Character Dictionary](/character-dictionary) — AniList character name matching and inline portraits
- [Jellyfin Integration](/jellyfin-integration) — remote playback and library connection
- [Jimaku Integration](/jimaku-integration) — subtitle fetching and API rate limits
- [YouTube Integration](/youtube-integration) — subtitle generation and playback
- [Immersion Tracking](/immersion-tracking) — telemetry and session logging
- [WebSocket / Texthooker API](/websocket-texthooker-api) — external texthooker clients
- [Subtitle Annotations](/subtitle-annotations) — N+1, frequency, JLPT, and name-match layers
- [Subtitle Sidebar](/subtitle-sidebar) — sidebar navigation and behavior
- [Configuration Reference](/configuration) — full config options
- [Shortcuts](/shortcuts) — keybinding reference
