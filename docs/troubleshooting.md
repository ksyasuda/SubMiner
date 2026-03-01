# Troubleshooting

Common issues and how to resolve them.

## MPV Connection

**Overlay starts but shows no subtitles**

SubMiner connects to mpv via a Unix socket (or named pipe on Windows). If the socket does not exist or the path does not match, the overlay will appear but subtitles will never arrive.

- Ensure mpv is running with `--input-ipc-server=/tmp/subminer-socket`.
- If you use a custom socket path, set it in both your mpv config and SubMiner config (`mpvSocketPath`).
- The `subminer` wrapper script sets the socket automatically when it launches mpv. If you launch mpv yourself, the `--input-ipc-server` flag is required.

SubMiner retries the connection automatically with increasing delays (200 ms, 500 ms, 1 s, 2 s on first connect; 1 s, 2 s, 5 s, 10 s on reconnect). If mpv exits and restarts, the overlay reconnects without needing a restart.

## Logging and App Mode

- Default log output is `info`.
- Use `--log-level` for more/less output.
- Use `--dev`/`--debug` only to force app/dev mode (for example to get dev behavior from the overlay/app); they do not change log verbosity.
- You can combine both, for example `SubMiner.AppImage --start --dev --log-level debug`, when you need maximum diagnostics.

## Performance and Resource Impact

### At a glance

- Baseline: `SubMiner --start` is usually lightweight for normal playback.
- Common spikes come from:
  - first subtitle parse/tokenization bursts
  - media generation (`ffmpeg` audio/image and AVIF paths)
  - media sync and subtitle tooling (`alass`, `ffsubsync`, `whisper` fallback path)
  - `ankiConnect` enrichment (plus polling overhead when proxy mode is disabled)

### If playback feels sluggish

1. Reduce overlay workload:

- set secondary subtitles hidden:
  - `secondarySub.defaultMode: "hidden"`
- disable optional enrichment:
  - `subtitleStyle.enableJlpt: false`
  - `subtitleStyle.frequencyDictionary.enabled: false`

2. Reduce rendering pressure:

- lower `subtitleStyle.fontSize`
- keep overlay complexity minimal during heavy CPU periods

3. Reduce media overhead:

- keep `ankiConnect.media.imageType` set to `static` (avoid animated AVIF unless needed)
- lower `ankiConnect.media.imageQuality`
- reduce `ankiConnect.media.maxMediaDuration`

4. Lower integration cost:

- disable AI translation when not needed (`ankiConnect.ai.enabled: false`)
- if needed, run immersion telemetry with lower duration expectations (`immersionTracking.enabled: false` for constrained sessions)
- prefer YouTube `--mode automatic` over `preprocess` on low-resource systems

### Practical low-impact profile

```json
{
  "subtitleStyle": {
    "fontSize": 30,
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
- Use `info` logs by default; keep `debug` for targeted diagnosis.
- Reproduce once with `SubMiner.AppImage --start --log-level debug` and open DevTools (`y` then `d`) if freezes recur.

**"Failed to parse MPV message"**

Logged when a malformed JSON line arrives from the mpv socket. Usually harmless — SubMiner skips the bad line and continues. If it happens constantly, check that nothing else is writing to the same socket path.

## AnkiConnect

**"AnkiConnect: unable to connect"**

SubMiner connects to the active Anki endpoint:

- `ankiConnect.url` (direct mode, default `http://127.0.0.1:8765`)
- `http://<ankiConnect.proxy.host>:<ankiConnect.proxy.port>` (proxy mode)

This error means the active endpoint is unavailable, or (in proxy mode) the proxy cannot reach `ankiConnect.proxy.upstreamUrl`.

- Install the [AnkiConnect](https://ankiweb.net/shared/info/2055492159) add-on in Anki.
- Make sure Anki is running before you start mining.
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
- On Linux, the overlay requires a compositor. Hyprland and Sway are supported natively; X11 requires `xdotool` and `xwininfo`.
- On macOS, grant Accessibility permission to SubMiner in System Settings > Privacy & Security > Accessibility.

**Overlay appears but clicks pass through / cannot interact**

- On Linux, mouse passthrough can be unreliable — this is a known Electron/platform limitation. The overlay keeps pointer events enabled by default on Linux.
- On macOS/Windows, `setIgnoreMouseEvents` toggles automatically. If clicks stop working, toggle the overlay off and back on (`Alt+Shift+O`).
- Make sure you are hovering over the subtitle area — the overlay only becomes interactive when the cursor is over subtitle text.

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

**"Yomitan extension not found in any search path"**

SubMiner bundles Yomitan and searches for it in these locations (in order):

1. `vendor/yomitan` (relative to executable)
2. `<resources>/yomitan` (Electron resources path)
3. `/usr/share/SubMiner/yomitan`
4. `~/.config/SubMiner/extensions/yomitan`

If you installed from the AppImage and see this error, the package may be incomplete. Re-download the AppImage or place the Yomitan extension manually in `~/.config/SubMiner/extensions/yomitan`.

**Yomitan popup does not appear when clicking words**

- Verify Yomitan loaded successfully — check the terminal output for "Loaded Yomitan extension".
- Yomitan requires dictionaries to be installed. Open Yomitan settings (`Alt+Shift+Y` or `SubMiner.AppImage --settings`) and confirm at least one dictionary is imported.
- If the overlay shows subtitles but words are not clickable, the tokenizer may have failed. See the MeCab section below.

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
- On Wayland, global shortcut registration has limitations depending on the compositor.

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

**"Configured alass executable not found"**

Install alass or configure the path:

- **Arch Linux (AUR)**: `yay -S alass-git`
- Set the path: `subsync.alass_path` in your config.

**"Subtitle synchronization failed"**

SubMiner tries alass first, then falls back to ffsubsync. If both fail:

- Ensure the reference subtitle track exists in the video (alass requires a source track).
- Check that `ffmpeg` is available (used to extract the internal subtitle track).
- Try running the sync tool manually to see detailed error output.

## Jimaku

**"Jimaku request failed" or HTTP 429**

The Jimaku API has rate limits. If you see 429 errors, wait for the retry duration shown in the OSD message and try again. If you have a Jimaku API key, set it in `jimaku.apiKey` or `jimaku.apiKeyCommand` to get higher rate limits.

## Platform-Specific

### Linux

- **Wayland (Hyprland/Sway)**: Window tracking uses compositor-specific commands. If `hyprctl` or `swaymsg` are not on `PATH`, tracking will fail silently.
- **X11**: Requires `xdotool` and `xwininfo`. If missing, the overlay cannot track the mpv window position.
- **Mouse passthrough**: On Linux, Electron's mouse passthrough is unreliable. SubMiner keeps pointer events enabled, meaning you may need to toggle the overlay off to interact with mpv controls underneath.

### macOS

- **Accessibility permission**: Required for window tracking. Grant it in System Settings > Privacy & Security > Accessibility.
- **Font rendering**: macOS uses a 0.87x font compensation factor for subtitle alignment between mpv and the overlay. If text alignment looks off, adjust subtitle offset by right-click dragging subtitle text.
- **Gatekeeper**: If macOS blocks SubMiner, right-click the app and select "Open" to bypass the warning, or remove the quarantine attribute: `xattr -d com.apple.quarantine /path/to/SubMiner.app`
