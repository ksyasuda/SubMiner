---
outline: [2, 3]
---

# Configuration

Settings are stored in `$XDG_CONFIG_HOME/SubMiner/config.jsonc` (or `~/.config/SubMiner/config.jsonc` when `XDG_CONFIG_HOME` is unset).

## Quick Start

For most users, start with this minimal configuration:

```json
{
  "ankiConnect": {
    "enabled": true,
    "deck": "YourDeckName",
    "fields": {
      "sentence": "Sentence",
      "audio": "Audio",
      "image": "Image"
    }
  }
}
```

Then customize as needed using the sections below.

## Configuration File

See [config.example.jsonc](/config.example.jsonc) for a comprehensive example configuration file with all available options, default values, and detailed comments. Only include the options you want to customize in your config file.

Generate a fresh default config from the centralized config registry:

```bash
SubMiner.AppImage --generate-config
SubMiner.AppImage --generate-config --config-path /tmp/subminer.jsonc
SubMiner.AppImage --generate-config --backup-overwrite
```

- `--generate-config` writes a default JSONC config template.
- JSONC config supports comments and trailing commas.
- If the target file exists, SubMiner prompts to create a timestamped backup and overwrite.
- In non-interactive shells, use `--backup-overwrite` to explicitly back up and overwrite.

Malformed config syntax (invalid JSON/JSONC) is startup-blocking: SubMiner shows a clear parse error with the config path and asks you to fix the file and restart.

For valid JSON/JSONC with invalid option values, SubMiner uses warn-and-fallback behavior: it logs the bad key/value and continues with the default for that option.

On macOS, these validation warnings also open a native dialog with full details (desktop notification banners can truncate long messages).

### Hot-Reload Behavior

SubMiner watches the active config file (`config.jsonc` or `config.json`) while running and applies supported updates automatically.

Hot-reloadable fields:

- `subtitleStyle`
- `keybindings`
- `shortcuts`
- `secondarySub.defaultMode`
- `ankiConnect.ai`

When these values change, SubMiner applies them live. Invalid config edits are rejected and the previous valid runtime config remains active.

Restart-required changes:

- Any other config sections still require restart.
- SubMiner shows an on-screen/system notification listing restart-required sections when they change.

### Configuration Options Overview

The configuration file includes several main sections:

**Core Settings**

- [**Logging**](#logging) - Runtime log level
- [**Auto-Start Overlay**](#auto-start-overlay) - Automatically show overlay on MPV connection
- [**Startup Warmups**](#startup-warmups) - Control what preloads on startup vs first-use defer
- [**WebSocket Server**](#websocket-server) - Built-in subtitle broadcasting server
- [**Texthooker**](#texthooker) - Control browser opening behavior

**Subtitle Display**

- [**Subtitle Style**](#subtitle-style) - Appearance customization
- [**Subtitle Position**](#subtitle-position) - Overlay vertical positioning
- [**Secondary Subtitles**](#secondary-subtitles) - Dual subtitle track support

**Keyboard & Controls**

- [**Keybindings**](#keybindings) - MPV command shortcuts
- [**Shortcuts Configuration**](#shortcuts-configuration) - Overlay keyboard shortcuts
- [**Manual Card Update Shortcuts**](#manual-card-update-shortcuts) - Shortcuts for manual Anki card workflows
- [**Session Help Modal**](#session-help-modal) - In-overlay shortcut reference
- [**Runtime Option Palette**](#runtime-option-palette) - Live, session-only option toggles

**Anki Integration**

- [**AnkiConnect**](#ankiconnect) - Automatic Anki card creation with media
- [**Kiku/Lapis Integration**](#kiku-lapis-integration) - Sentence cards and duplicate handling for Kiku/Lapis note types
- [**N+1 Word Highlighting**](#n1-word-highlighting) - Known-word cache and single-target highlighting
- [**Field Grouping Modes**](#field-grouping-modes) - Kiku/Lapis duplicate card merging

**External Integrations**

- [**Jimaku**](#jimaku) - Jimaku API configuration and defaults
- [**Auto Subtitle Sync**](#auto-subtitle-sync) - Sync current subtitle with `alass`/`ffsubsync`
- [**AniList**](#anilist) - Optional post-watch progress updates
- [**Jellyfin**](#jellyfin) - Optional Jellyfin auth, library listing, and playback launch
- [**Discord Rich Presence**](#discord-rich-presence) - Optional Discord activity card updates
- [**Immersion Tracking**](#immersion-tracking) - Track subtitle sessions and mining activity in SQLite
- [**YouTube Subtitle Generation**](#youtube-subtitle-generation) - Launcher defaults for yt-dlp + local whisper fallback

## Core Settings

### Logging

Control the minimum log level for runtime output:

```json
{
  "logging": {
    "level": "info"
  }
}
```

| Option  | Values                              | Description                                      |
| ------- | ----------------------------------- | ------------------------------------------------ |
| `level` | `"debug"`, `"info"`, `"warn"`, `"error"` | Minimum log level for runtime logging (default: `"info"`) |

### Auto-Start Overlay

Control whether the overlay automatically becomes visible when it connects to mpv:

```json
{
  "auto_start_overlay": false
}
```

| Option               | Values          | Description                                            |
| -------------------- | --------------- | ------------------------------------------------------ |
| `auto_start_overlay` | `true`, `false` | Auto-show overlay on mpv connection (default: `false`) |

The mpv plugin controls startup overlay visibility via `auto_start_visible_overlay` in `subminer.conf`.
For wrapper-driven playback, `subminer.conf` can also enable startup pause gating with
`auto_start_pause_until_ready` (requires `auto_start=yes` + `auto_start_visible_overlay=yes`).
Current plugin defaults in `subminer.conf` are:

- `auto_start=yes`
- `auto_start_visible_overlay=yes`
- `auto_start_pause_until_ready=yes`

### Startup Warmups

Control which startup warmups run in the background versus deferring to first real usage:

```json
{
  "startupWarmups": {
    "lowPowerMode": false,
    "mecab": true,
    "yomitanExtension": true,
    "subtitleDictionaries": true,
    "jellyfinRemoteSession": true
  }
}
```

| Option                  | Values          | Description                                                                                       |
| ----------------------- | --------------- | ------------------------------------------------------------------------------------------------- |
| `lowPowerMode`          | `true`, `false` | Defer all warmups except Yomitan extension                                                        |
| `mecab`                 | `true`, `false` | Warm up MeCab tokenizer at startup                                                                |
| `yomitanExtension`      | `true`, `false` | Warm up Yomitan extension at startup                                                              |
| `subtitleDictionaries`  | `true`, `false` | Warm up JLPT + frequency dictionaries at startup                                                  |
| `jellyfinRemoteSession` | `true`, `false` | Warm up Jellyfin remote session at startup (still requires Jellyfin remote auto-connect settings) |

Defaults warm everything (`true` for all toggles, `lowPowerMode: false`). Setting a warmup toggle to `false` defers that work until first usage.

### WebSocket Server

The overlay includes a built-in WebSocket server that broadcasts subtitle text to connected clients (such as texthooker-ui) for external processing.

By default, the server uses "auto" mode: it starts automatically unless [mpv_websocket](https://github.com/kuroahna/mpv_websocket) is detected at `~/.config/mpv/mpv_websocket`. If you have mpv_websocket installed, the built-in server is skipped to avoid conflicts.

See `config.example.jsonc` for detailed configuration options.

```json
{
  "websocket": {
    "enabled": "auto",
    "port": 6677
  }
}
```

| Option    | Values                    | Description                                              |
| --------- | ------------------------- | -------------------------------------------------------- |
| `enabled` | `true`, `false`, `"auto"` | `"auto"` (default) disables if mpv_websocket is detected |
| `port`    | number                    | WebSocket server port (default: 6677)                    |

### Texthooker

Control whether the browser opens automatically when texthooker starts:

See `config.example.jsonc` for detailed configuration options.

```json
{
  "texthooker": {
    "openBrowser": true
  }
}
```

## Subtitle Display

### Subtitle Style

Customize the appearance of primary and secondary subtitles:

See `config.example.jsonc` for detailed configuration options.

```json
{
  "subtitleStyle": {
    "fontFamily": "M PLUS 1 Medium, Source Han Sans JP, Noto Sans CJK JP",
    "fontSize": 35,
    "fontColor": "#cad3f5",
    "fontWeight": "600",
    "lineHeight": 1.35,
    "letterSpacing": "-0.01em",
    "wordSpacing": 0,
    "fontKerning": "normal",
    "textRendering": "geometricPrecision",
    "textShadow": "0 3px 10px rgba(0,0,0,0.69)",
    "fontStyle": "normal",
    "backgroundColor": "rgb(30, 32, 48, 0.88)",
    "backdropFilter": "blur(6px)",
    "secondary": {
      "fontFamily": "Inter, Noto Sans, Helvetica Neue, sans-serif",
      "fontSize": 24,
      "fontColor": "#cad3f5",
      "backgroundColor": "transparent"
    }
  }
}
```

| Option                             | Values      | Description                                                                                                                |
| ---------------------------------- | ----------- | -------------------------------------------------------------------------------------------------------------------------- |
| `fontFamily`                       | string      | CSS font-family value (default: `"M PLUS 1 Medium, Source Han Sans JP, Noto Sans CJK JP"`)                                 |
| `fontSize`                         | number (px) | Font size in pixels (default: `35`)                                                                                        |
| `fontColor`                        | string      | Any CSS color value (default: `"#cad3f5"`)                                                                                 |
| `fontWeight`                       | string      | CSS font-weight, e.g. `"bold"`, `"normal"`, `"600"` (default: `"600"`)                                                     |
| `fontStyle`                        | string      | `"normal"` or `"italic"` (default: `"normal"`)                                                                             |
| `backgroundColor`                  | string      | Any CSS color, including `"transparent"` (default: `"rgb(30, 32, 48, 0.88)"`)                                              |
| `enableJlpt`                       | boolean     | Enable JLPT level underline styling (`false` by default)                                                                   |
| `preserveLineBreaks`               | boolean     | Preserve line breaks in visible overlay subtitle rendering (`false` by default). Enable to mirror mpv line layout.         |
| `autoPauseVideoOnHover`            | boolean     | Pause playback while mouse hovers subtitle text, then resume on leave (`true` by default).                                  |
| `hoverTokenColor`                  | string      | Hex color used for hovered subtitle token highlight in mpv (default: catppuccin mauve)                                     |
| `hoverTokenBackgroundColor`        | string      | CSS color used for hovered subtitle token background highlight (default: semi-transparent dark)                            |
| `frequencyDictionary.enabled`      | boolean     | Enable frequency highlighting from dictionary lookups (`false` by default)                                                 |
| `frequencyDictionary.sourcePath`   | string      | Path to a local frequency dictionary root. Leave empty or omit to use installed/default frequency-dictionary search paths. |
| `frequencyDictionary.topX`         | number      | Only color tokens whose frequency rank is `<= topX` (`1000` by default)                                                    |
| `frequencyDictionary.mode`         | string      | `"single"` or `"banded"` (`"single"` by default)                                                                           |
| `frequencyDictionary.matchMode`    | string      | `"headword"` or `"surface"` (`"headword"` by default)                                                                      |
| `frequencyDictionary.singleColor`  | string      | Color used for all highlighted tokens in single mode                                                                       |
| `frequencyDictionary.bandedColors` | string[]    | Array of five hex colors used for ranked bands in banded mode                                                              |
| `nPlusOneColor`                    | string      | Existing n+1 highlight color (default: `#c6a0f6`)                                                                          |
| `knownWordColor`                   | string      | Existing known-word highlight color (default: `#a6da95`)                                                                   |
| `jlptColors`                       | object      | JLPT level underline colors object (`N1`..`N5`)                                                                            |
| `secondary`                        | object      | Override any of the above for secondary subtitles (optional)                                                               |

JLPT underlining is powered by offline term-meta bank files at runtime. See [`docs/jlpt-vocab-bundle.md`](jlpt-vocab-bundle.md) for required files, source/version refresh steps, and deterministic fallback behavior.

Frequency dictionary highlighting uses the same dictionary file format as JLPT bundle lookups (`term_meta_bank_*.json` under discovered dictionary directories). A token is highlighted when it has a positive integer `frequencyRank` (lower is more common) and the rank is within `topX`.

Lookup behavior:

- Set `frequencyDictionary.sourcePath` to a directory containing `term_meta_bank_*.json` for a fully custom source.
- If `sourcePath` is missing or empty, SubMiner searches default install/runtime locations for `frequency-dictionary` directories (for example app resources, user data paths, and current working directory).
- In both cases, only terms with a valid `frequencyRank` are used; everything else falls back to no highlighting.
- `frequencyDictionary.matchMode` controls which token text is used for frequency lookups: `headword` (dictionary form) or `surface` (visible subtitle text).
- Frequency highlighting skips tokens that look like non-lexical SFX/interjection noise (for example kana reduplication or short kana endings like `っ`), even when dictionary ranks exist.

In `single` mode all highlights use `singleColor`; in `banded` mode tokens map to five ascending color bands from most common to least common inside the topX window.

Secondary subtitle defaults: `fontFamily: "Inter, Noto Sans, Helvetica Neue, sans-serif"`, `fontSize: 24`, `fontColor: "#cad3f5"`, `backgroundColor: "transparent"`. Any property not set in `secondary` falls back to the CSS defaults.

**See `config.example.jsonc`** for the complete list of subtitle style configuration options.

`jlptColors` keys are:

| Key  | Default   | Description             |
| ---- | --------- | ----------------------- |
| `N1` | `#ed8796` | JLPT N1 underline color |
| `N2` | `#f5a97f` | JLPT N2 underline color |
| `N3` | `#f9e2af` | JLPT N3 underline color |
| `N4` | `#a6e3a1` | JLPT N4 underline color |
| `N5` | `#8aadf4` | JLPT N5 underline color |

**Image Quality Notes:**

- `imageQuality` affects JPG and WebP only; PNG is lossless and ignores this setting
- JPG quality is mapped to FFmpeg's scale (2-31, lower = better)
- WebP quality uses FFmpeg's native 0-100 scale

### Subtitle Position

Set the initial vertical subtitle position (measured from the bottom of the screen):

```json
{
  "subtitlePosition": {
    "yPercent": 10
  }
}
```

| Option     | Values           | Description                                                            |
| ---------- | ---------------- | ---------------------------------------------------------------------- |
| `yPercent` | number (0 - 100) | Distance from the bottom as a percent of screen height (default: `10`) |
In the overlay, you can fine-tune subtitle position at runtime with `Right-click + drag` on subtitle text.

### Secondary Subtitles

Display a second subtitle track (e.g., English alongside Japanese) in the overlay:

See `config.example.jsonc` for detailed configuration options.

```json
{
  "secondarySub": {
    "secondarySubLanguages": ["eng", "en"],
    "autoLoadSecondarySub": true,
    "defaultMode": "hover"
  }
}
```

| Option                  | Values                             | Description                                            |
| ----------------------- | ---------------------------------- | ------------------------------------------------------ |
| `secondarySubLanguages` | string[]                           | Language codes to auto-load (e.g., `["eng", "en"]`)    |
| `autoLoadSecondarySub`  | `true`, `false`                    | Auto-detect and load matching secondary subtitle track |
| `defaultMode`           | `"hidden"`, `"visible"`, `"hover"` | Initial display mode (default: `"hover"`)              |

**Display modes:**

- **hidden** — Secondary subtitles not shown
- **visible** — Always visible at top of overlay
- **hover** — Only visible when hovering over the subtitle area (default)

**See `config.example.jsonc`** for additional secondary subtitle configuration options.

## Keyboard & Controls

### Keybindings

Add a `keybindings` array to configure keyboard shortcuts that send commands to mpv:

See `config.example.jsonc` for detailed configuration options and more examples.

**Default keybindings:**

| Key               | Command                      | Description                           |
| ----------------- | ---------------------------- | ------------------------------------- |
| `Space`           | `["cycle", "pause"]`         | Toggle pause                          |
| `KeyJ`            | `["cycle", "sid"]`           | Cycle primary subtitle track          |
| `Shift+KeyJ`      | `["cycle", "secondary-sid"]` | Cycle secondary subtitle track        |
| `ArrowRight`      | `["seek", 5]`                | Seek forward 5 seconds                |
| `ArrowLeft`       | `["seek", -5]`               | Seek backward 5 seconds               |
| `ArrowUp`         | `["seek", 60]`               | Seek forward 60 seconds               |
| `ArrowDown`       | `["seek", -60]`              | Seek backward 60 seconds              |
| `Shift+KeyH`      | `["sub-seek", -1]`           | Jump to previous subtitle             |
| `Shift+KeyL`      | `["sub-seek", 1]`            | Jump to next subtitle                 |
| `Ctrl+Shift+KeyH` | `["__replay-subtitle"]`      | Replay current subtitle, pause at end |
| `Ctrl+Shift+KeyL` | `["__play-next-subtitle"]`   | Play next subtitle, pause at end      |
| `KeyQ`            | `["quit"]`                   | Quit mpv                              |
| `Ctrl+KeyW`       | `["quit"]`                   | Quit mpv                              |

**Custom keybindings example:**

```json
{
  "keybindings": [
    { "key": "ArrowRight", "command": ["seek", 5] },
    { "key": "ArrowLeft", "command": ["seek", -5] },
    { "key": "Shift+ArrowRight", "command": ["seek", 30] },
    { "key": "KeyR", "command": ["script-binding", "immersive/auto-replay"] },
    { "key": "KeyA", "command": ["script-message", "ankiconnect-add-note"] }
  ]
}
```

**Key format:** Use `KeyboardEvent.code` values (`Space`, `ArrowRight`, `KeyR`, etc.) with optional modifiers (`Ctrl+`, `Alt+`, `Shift+`, `Meta+`).

**Disable a default binding:** Set command to `null`:

```json
{ "key": "Space", "command": null }
```

**Special commands:** Commands prefixed with `__` are handled internally by the overlay rather than sent to mpv. `__replay-subtitle` replays the current subtitle and pauses at its end. `__play-next-subtitle` seeks to the next subtitle, plays it, and pauses at its end. `__runtime-options-open` opens the runtime options palette. `__runtime-option-cycle:<id>[:next|prev]` cycles a runtime option value.

**Supported commands:** Any valid mpv JSON IPC command array (`["cycle", "pause"]`, `["seek", 5]`, `["script-binding", "..."]`, etc.)

For subtitle-position and subtitle-track proxy commands (`sub-pos`, `sid`, `secondary-sid`), SubMiner also shows an mpv OSD notification after the command runs.

**See `config.example.jsonc`** for more keybinding examples and configuration options.

### Shortcuts Configuration

Customize or disable the overlay keyboard shortcuts:

See `config.example.jsonc` for detailed configuration options.

```json
{
  "shortcuts": {
    "toggleVisibleOverlayGlobal": "Alt+Shift+O",
    "copySubtitle": "CommandOrControl+C",
    "copySubtitleMultiple": "CommandOrControl+Shift+C",
    "updateLastCardFromClipboard": "CommandOrControl+V",
    "triggerFieldGrouping": "CommandOrControl+G",
    "triggerSubsync": "Ctrl+Alt+S",
    "mineSentence": "CommandOrControl+S",
    "mineSentenceMultiple": "CommandOrControl+Shift+S",
    "markAudioCard": "CommandOrControl+Shift+A",
    "openRuntimeOptions": "CommandOrControl+Shift+O",
    "openJimaku": "Ctrl+Shift+J",
    "multiCopyTimeoutMs": 3000
  }
}
```

| Option                        | Values           | Description                                                                                                                                   |
| ----------------------------- | ---------------- | --------------------------------------------------------------------------------------------------------------------------------------------- |
| `toggleVisibleOverlayGlobal`  | string \| `null` | Global accelerator for toggling visible subtitle overlay (default: `"Alt+Shift+O"`)                                                           |
| `copySubtitle`                | string \| `null` | Accelerator for copying current subtitle (default: `"CommandOrControl+C"`)                                                                    |
| `copySubtitleMultiple`        | string \| `null` | Accelerator for multi-copy mode (default: `"CommandOrControl+Shift+C"`)                                                                       |
| `updateLastCardFromClipboard` | string \| `null` | Accelerator for updating card from clipboard (default: `"CommandOrControl+V"`)                                                                |
| `triggerFieldGrouping`        | string \| `null` | Accelerator for Kiku field grouping on last card (default: `"CommandOrControl+G"`; only active when `behavior.autoUpdateNewCards` is `false`) |
| `triggerSubsync`              | string \| `null` | Accelerator for running Subsync (default: `"Ctrl+Alt+S"`)                                                                                     |
| `mineSentence`                | string \| `null` | Accelerator for creating sentence card from current subtitle (default: `"CommandOrControl+S"`)                                                |
| `mineSentenceMultiple`        | string \| `null` | Accelerator for multi-mine sentence card mode (default: `"CommandOrControl+Shift+S"`)                                                         |
| `multiCopyTimeoutMs`          | number           | Timeout in ms for multi-copy/mine digit input (default: `3000`)                                                                               |
| `toggleSecondarySub`          | string \| `null` | Accelerator for cycling secondary subtitle mode (default: `"CommandOrControl+Shift+V"`)                                                       |
| `markAudioCard`               | string \| `null` | Accelerator for marking last card as audio card (default: `"CommandOrControl+Shift+A"`)                                                       |
| `openRuntimeOptions`          | string \| `null` | Opens runtime options palette for live session-only toggles (default: `"CommandOrControl+Shift+O"`)                                           |
| `openJimaku`                  | string \| `null` | Opens the Jimaku search modal (default: `"Ctrl+Shift+J"`)                                                                                     |

**See `config.example.jsonc`** for the complete list of shortcut configuration options.

Set any shortcut to `null` to disable it.

Feature-dependent shortcuts/keybindings only run when their related integration is enabled. For example, Anki/Kiku shortcuts require `ankiConnect.enabled` (and Kiku-specific behavior where applicable), and Jellyfin remote startup behavior requires Jellyfin to be enabled.

### Manual Card Update Shortcuts

When `behavior.autoUpdateNewCards` is set to `false`, new cards are detected but not automatically updated. Use these keyboard shortcuts for manual control:

| Shortcut       | Action                                                                                                             |
| -------------- | ------------------------------------------------------------------------------------------------------------------ |
| `Ctrl+C`       | Copy the current subtitle line to clipboard (preserves line breaks)                                                |
| `Ctrl+Shift+C` | Enter multi-copy mode. Press `1-9` to copy that many recent lines, or `Esc` to cancel. Timeout: 3 seconds          |
| `Ctrl+V`       | Update the last added Anki card using subtitles from clipboard                                                     |
| `Ctrl+G`       | Trigger Kiku duplicate field grouping for the last added card (only when `behavior.autoUpdateNewCards` is `false`) |
| `Ctrl+S`       | Create a sentence card from the current subtitle line                                                              |
| `Ctrl+Shift+S` | Enter multi-mine mode. Press `1-9` to create a sentence card from that many recent lines, or `Esc` to cancel       |
| `Ctrl+Shift+V` | Cycle secondary subtitle display mode (hidden → visible → hover)                                                   |
| `Ctrl+Shift+A` | Mark the last added Anki card as an audio card (sets IsAudioCard, SentenceAudio, Sentence, Picture)                |
| `Ctrl+Shift+O` | Open runtime options palette (session-only live toggles)                                                           |
| `Ctrl/Cmd+A`   | Append clipboard video path to MPV playlist (fixed, not currently configurable)                                    |

**Multi-line copy workflow:**

1. Press `Ctrl+Shift+C`
2. Press a number key (`1-9`) within 3 seconds
3. The specified number of most recent subtitle lines are copied
4. Press `Ctrl+V` to update the last added card with the copied lines

These shortcuts are only active when the overlay window is visible and automatically disabled when hidden.

### Session Help Modal

The session help modal is opened with `Y-H` by default (falls back to `Y-K` if needed) and shows the current session keybindings and color legend.

You can filter the modal quickly with `/`:

- Type any part of the action name or shortcut in the search bar.
- Search is case-insensitive and ignores spaces/punctuation (`+`, `-`, `_`, `/`) so `ctrl w`, `ctrl+w`, and `ctrl+s` all match.
- Results are filtered across active MPV shortcuts, configured overlay shortcuts, and color legend items.

While the modal is open:

- `Esc`: close the modal (or clear the filter when text is entered)
- `↑/↓`, `j/k`: move selection
- Mouse/trackpad: click to select and activate rows

The list is generated at runtime from:

- Your active mpv keybindings (`keybindings`).
- Your configured overlay shortcuts (`shortcuts`, including runtime-loaded config values).
- Current subtitle color settings from `subtitleStyle`.

When config hot-reload updates shortcut/keybinding/style values, close and reopen the help modal to refresh the displayed entries.

### Runtime Option Palette

Use the runtime options palette to toggle settings live while SubMiner is running. These changes are session-only and reset on restart.

Current runtime options:

- `ankiConnect.behavior.autoUpdateNewCards` (`On` / `Off`)
- `ankiConnect.nPlusOne.highlightEnabled` (`On` / `Off`)
- `subtitleStyle.enableJlpt` (`On` / `Off`)
- `subtitleStyle.frequencyDictionary.enabled` (`On` / `Off`)
- `ankiConnect.nPlusOne.matchMode` (`headword` / `surface`)
- `ankiConnect.isKiku.fieldGrouping` (`auto` / `manual` / `disabled`)

Annotation toggles (`nPlusOne`, `enableJlpt`, `frequencyDictionary.enabled`) only apply to new subtitle lines after the toggle. The currently displayed line is not re-tokenized in place.

Default shortcut: `Ctrl+Shift+O`

Palette controls:

- `Arrow Up/Down`: select option
- `Arrow Left/Right`: change selected value
- `Enter`: apply selected value
- `Esc`: close

## Anki Integration

### AnkiConnect

Enable automatic Anki card creation and updates with media generation:

```json
{
  "ankiConnect": {
    "enabled": true,
    "url": "http://127.0.0.1:8765",
    "pollingRate": 3000,
    "proxy": {
      "enabled": false,
      "host": "127.0.0.1",
      "port": 8766,
      "upstreamUrl": "http://127.0.0.1:8765"
    },
    "tags": ["SubMiner"],
    "deck": "Learning::Japanese",
    "fields": {
      "audio": "ExpressionAudio",
      "image": "Picture",
      "sentence": "Sentence",
      "miscInfo": "MiscInfo",
      "translation": "SelectionText"
    },
    "ai": {
      "enabled": false,
      "alwaysUseAiTranslation": false,
      "apiKey": "",
      "model": "openai/gpt-4o-mini",
      "baseUrl": "https://openrouter.ai/api",
      "targetLanguage": "English",
      "systemPrompt": "You are a translation engine. Return only the translated text with no explanations."
    },
    "media": {
      "generateAudio": true,
      "generateImage": true,
      "imageType": "static",
      "imageFormat": "jpg",
      "imageQuality": 92,
      "imageMaxWidth": 1280,
      "imageMaxHeight": 720,
      "animatedFps": 10,
      "animatedMaxWidth": 640,
      "animatedMaxHeight": 360,
      "animatedCrf": 35,
      "audioPadding": 0.5,
      "fallbackDuration": 3,
      "maxMediaDuration": 30
    },
    "behavior": {
      "autoUpdateNewCards": true,
      "overwriteAudio": true,
      "overwriteImage": true
    },
    "metadata": {
      "pattern": "[SubMiner] %f (%t)"
    },
    "isLapis": {
      "enabled": true,
      "sentenceCardModel": "Japanese sentences"
    },
    "isKiku": {
      "enabled": false,
      "fieldGrouping": "disabled",
      "deleteDuplicateInAuto": true
    }
  }
}
```

This example is intentionally compact. The option table below documents available `ankiConnect` settings and behavior.

**Requirements:** [AnkiConnect](https://github.com/FooSoft/anki-connect) plugin must be installed and running in Anki. ffmpeg must be installed for media generation.

| Option                                  | Values                                  | Description                                                                                                                                   |
| --------------------------------------- | --------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------- |
| `enabled`                               | `true`, `false`                         | Enable AnkiConnect integration (default: `false`)                                                                                             |
| `url`                                   | string (URL)                            | AnkiConnect API URL (default: `http://127.0.0.1:8765`)                                                                                        |
| `pollingRate`                           | number (ms)                             | How often to check for new cards in polling mode (default: `3000`; ignored for direct proxy `addNote`/`addNotes` updates)                     |
| `proxy.enabled`                         | `true`, `false`                         | Enable local AnkiConnect-compatible proxy for push-based auto-enrichment (default: `true`)                                                    |
| `proxy.host`                            | string                                  | Bind host for local AnkiConnect proxy (default: `127.0.0.1`)                                                                                  |
| `proxy.port`                            | number                                  | Bind port for local AnkiConnect proxy (default: `8766`)                                                                                       |
| `proxy.upstreamUrl`                     | string (URL)                            | Upstream AnkiConnect URL that proxy forwards to (default: `http://127.0.0.1:8765`)                                                            |
| `tags`                                  | array of strings                        | Tags automatically added to cards mined/updated by SubMiner (default: `['SubMiner']`; set `[]` to disable automatic tagging).                 |
| `deck`                                  | string                                  | Anki deck to monitor for new cards                                                                                                            |
| `ankiConnect.nPlusOne.decks`            | array of strings                        | Decks used for N+1 known-word cache lookups. When omitted/empty, falls back to `ankiConnect.deck`.                                            |
| `fields.audio`                          | string                                  | Card field for audio files (default: `ExpressionAudio`)                                                                                       |
| `fields.image`                          | string                                  | Card field for images (default: `Picture`)                                                                                                    |
| `fields.sentence`                       | string                                  | Card field for sentences (default: `Sentence`)                                                                                                |
| `fields.miscInfo`                       | string                                  | Card field for metadata (default: `"MiscInfo"`, set to `null` to disable)                                                                     |
| `fields.translation`                    | string                                  | Card field for sentence-card translation/back text (default: `SelectionText`)                                                                 |
| `ai.enabled`                            | `true`, `false`                         | Use AI translation for sentence cards. Also auto-attempted when secondary subtitle is missing.                                                |
| `ai.alwaysUseAiTranslation`             | `true`, `false`                         | When `true`, always use AI translation even if secondary subtitles exist. When `false`, AI is used only when no secondary subtitle exists.    |
| `ai.apiKey`                             | string                                  | API key for your OpenAI-compatible endpoint (required for translation).                                                                       |
| `ai.model`                              | string                                  | Model id for your OpenAI-compatible endpoint (default: `openai/gpt-4o-mini`).                                                                 |
| `ai.baseUrl`                            | string (URL)                            | OpenAI-compatible API base URL; accepts with or without `/v1`.                                                                                |
| `ai.targetLanguage`                     | string                                  | Target language name used in translation prompt (default: `English`).                                                                         |
| `ai.systemPrompt`                       | string                                  | System prompt used for translation (default returns translation text only).                                                                   |
| `media.generateAudio`                   | `true`, `false`                         | Generate audio clips from video (default: `true`)                                                                                             |
| `media.generateImage`                   | `true`, `false`                         | Generate image/animation screenshots (default: `true`)                                                                                        |
| `media.imageType`                       | `"static"`, `"avif"`                    | Image type: static screenshot or animated AVIF (default: `"static"`)                                                                          |
| `media.imageFormat`                     | `"jpg"`, `"png"`, `"webp"`              | Image format (default: `"jpg"`)                                                                                                               |
| `media.imageQuality`                    | number (1-100)                          | Image quality for JPG/WebP; PNG ignores this (default: `92`)                                                                                  |
| `media.imageMaxWidth`                   | number (px)                             | Optional max width for static screenshots. Unset keeps source width.                                                                          |
| `media.imageMaxHeight`                  | number (px)                             | Optional max height for static screenshots. Unset keeps source height.                                                                        |
| `media.animatedFps`                     | number (1-60)                           | FPS for animated AVIF (default: `10`)                                                                                                         |
| `media.animatedMaxWidth`                | number (px)                             | Max width for animated AVIF (default: `640`)                                                                                                  |
| `media.animatedMaxHeight`               | number (px)                             | Optional max height for animated AVIF. Unset keeps source aspect-constrained height.                                                          |
| `media.animatedCrf`                     | number (0-63)                           | CRF quality for AVIF; lower = higher quality (default: `35`)                                                                                  |
| `media.audioPadding`                    | number (seconds)                        | Padding around audio clip timing (default: `0.5`)                                                                                             |
| `media.fallbackDuration`                | number (seconds)                        | Default duration if timing unavailable (default: `3.0`)                                                                                       |
| `media.maxMediaDuration`                | number (seconds)                        | Max duration for generated media from multi-line copy (default: `30`, `0` to disable)                                                         |
| `behavior.overwriteAudio`               | `true`, `false`                         | Replace existing audio on updates; when `false`, new audio is appended/prepended per `behavior.mediaInsertMode` (default: `true`)             |
| `behavior.overwriteImage`               | `true`, `false`                         | Replace existing images on updates; when `false`, new images are appended/prepended per `behavior.mediaInsertMode` (default: `true`)          |
| `behavior.mediaInsertMode`              | `"append"`, `"prepend"`                 | Where to insert new media when overwrite is off (default: `"append"`)                                                                         |
| `behavior.highlightWord`                | `true`, `false`                         | Highlight the word in sentence context (default: `true`)                                                                                      |
| `ankiConnect.nPlusOne.highlightEnabled` | `true`, `false`                         | Enable fast local highlighting for words already known in Anki (default: `false`)                                                             |
| `ankiConnect.nPlusOne.nPlusOne`         | hex color string                        | Text color for the single target token to study when exactly one unknown candidate exists in a sentence (default: `"#c6a0f6"`).               |
| `ankiConnect.nPlusOne.knownWord`        | hex color string                        | Legacy known-word color kept for backward compatibility (default: `"#a6da95"`).                                                               |
| `ankiConnect.nPlusOne.matchMode`        | `"headword"`, `"surface"`               | Matching strategy for known-word highlighting (default: `"headword"`). `headword` uses token headwords; `surface` uses visible subtitle text. |
| `ankiConnect.nPlusOne.minSentenceWords` | number                                  | Minimum number of words required in a sentence before single unknown-word N+1 highlighting can trigger (default: `3`).                        |
| `ankiConnect.nPlusOne.refreshMinutes`   | number                                  | Minutes between known-word cache refreshes (default: `1440`)                                                                                  |
| `ankiConnect.nPlusOne.decks`            | array of strings                        | Decks used by known-word cache refresh. Leave empty for compatibility with legacy `deck` scope.                                               |
| `behavior.notificationType`             | `"osd"`, `"system"`, `"both"`, `"none"` | Notification type on card update (default: `"osd"`)                                                                                           |
| `behavior.autoUpdateNewCards`           | `true`, `false`                         | Automatically update cards on creation (default: `true`)                                                                                      |
| `metadata.pattern`                      | string                                  | Format pattern for metadata: `%f`=filename, `%F`=filename+ext, `%t`=time                                                                      |
| `isLapis`                               | object                                  | Lapis/shared sentence-card config: `{ enabled, sentenceCardModel }`. Sentence/audio field names are fixed to `Sentence` and `SentenceAudio`.  |
| `isKiku`                                | object                                  | Kiku-only config: `{ enabled, fieldGrouping, deleteDuplicateInAuto }` (shared sentence/audio/model settings are inherited from `isLapis`)     |

### Kiku/Lapis Integration

SubMiner is intentionally built for [Kiku](https://kiku.youyoumu.my.id/) and [Lapis](https://github.com/donkuri/lapis) workflows, with note-type-specific behavior built into Anki settings.

```jsonc
"ankiConnect": {
  "isLapis": {
    "enabled": true,
    "sentenceCardModel": "Japanese sentences"
  },
  "isKiku": {
    "enabled": true,
    "fieldGrouping": "manual",
    "deleteDuplicateInAuto": true
  }
}
```

- Enable `isLapis` to mine dedicated sentence cards. SubMiner sets `IsSentenceCard` to `"x"` and fills the sentence fields for the configured model.
- Enable `isKiku` to turn on duplicate merge behavior for mined Word/Expression hits.
- When both are enabled, Kiku behavior is applied for grouping while sentence-card model settings are still read from `isLapis`.
- `isKiku.fieldGrouping` supports `disabled`, `auto`, and `manual` merge modes; see [Field Grouping Modes](#field-grouping-modes).

### N+1 Word Highlighting

When `ankiConnect.nPlusOne.highlightEnabled` is enabled, SubMiner builds a local cache of known words from Anki to highlight already learned tokens in subtitle rendering.

Known-word cache policy:

- Initial sync runs when the integration starts if the cache is missing or stale.
- `ankiConnect.nPlusOne.refreshMinutes` controls the minimum time between refreshes; between refreshes, cached words are reused without querying Anki.
- `ankiConnect.nPlusOne.nPlusOne` sets the color for the single target token when exactly one eligible unknown word exists.
- `ankiConnect.nPlusOne.minSentenceWords` sets the minimum token count required in a sentence for N+1 highlighting (default: `3`).
- `ankiConnect.nPlusOne.knownWord` sets the legacy known-word highlight color for tokens already in Anki.
- `ankiConnect.nPlusOne.decks` accepts one or more decks. If empty, it uses the legacy single `ankiConnect.deck` value as scope.
- Cache state is persisted to `known-words-cache.json` under the app `userData` directory.
- The cache is automatically invalidated when the configured scope changes (for example, when deck changes).
- Cache lookups are in-memory. By default, token headwords are matched against cached `Expression` / `Word` values; set `ankiConnect.nPlusOne.matchMode` to `"surface"` for raw subtitle text matching.
- `ankiConnect.behavior.nPlusOne*` legacy keys (`nPlusOneHighlightEnabled`, `nPlusOneRefreshMinutes`, `nPlusOneMatchMode`) are deprecated and only kept for backward compatibility.
- Legacy top-level `ankiConnect` migration keys (for example `audioField`, `generateAudio`, `imageType`) are compatibility-only, validated before mapping, and ignored with a warning when invalid.
- If AnkiConnect is unreachable, the cache remains in its previous state and an on-screen/system status message is shown.
- Known-word sync activity is logged at `INFO`/`DEBUG` level with the `anki` logger scope and includes scope, notes returned, and word counts.

To refresh roughly once per day, set:

```json
{
  "ankiConnect": {
    "nPlusOne": {
      "highlightEnabled": true,
      "refreshMinutes": 1440
    }
  }
}
```

### Field Grouping Modes

| Mode       | Behavior                                                                                                                   |
| ---------- | -------------------------------------------------------------------------------------------------------------------------- |
| `auto`     | Automatically merges the new card's content into the original; duplicate deletion is controlled by `deleteDuplicateInAuto` |
| `manual`   | Shows an overlay popup to choose which card to keep and whether to delete the duplicate after merge                        |
| `disabled` | No field grouping; duplicate cards are left as-is                                                                          |

`deleteDuplicateInAuto` controls whether `auto` mode deletes the duplicate after merge (default: `true`). In `manual` mode, the popup asks each time whether to delete the duplicate.

<video controls playsinline preload="metadata" poster="/assets/kiku-integration-poster.jpg" style="width: 100%; max-width: 960px;">
  <source :src="'/assets/kiku-integration.webm'" type="video/webm" />
  Your browser does not support the video tag.
</video>

<a :href="'/assets/kiku-integration.webm'" target="_blank" rel="noreferrer">Open demo in a new tab</a>

## External Integrations

### Jimaku

Configure Jimaku API access and defaults:

```json
{
  "jimaku": {
    "apiKey": "YOUR_API_KEY",
    "apiKeyCommand": "cat ~/.jimaku_key",
    "apiBaseUrl": "https://jimaku.cc",
    "languagePreference": "ja",
    "maxEntryResults": 10
  }
}
```

Jimaku is rate limited; if you hit a limit, SubMiner will surface the retry delay from the API response.

Set `openBrowser` to `false` to only print the URL without opening a browser.

### Auto Subtitle Sync

Sync the active subtitle track using `alass` (preferred) or `ffsubsync`:

```json
{
  "subsync": {
    "defaultMode": "auto",
    "alass_path": "",
    "ffsubsync_path": "",
    "ffmpeg_path": ""
  }
}
```

| Option           | Values               | Description                                                                                                 |
| ---------------- | -------------------- | ----------------------------------------------------------------------------------------------------------- |
| `defaultMode`    | `"auto"`, `"manual"` | `auto`: try `alass` against secondary subtitle, then fallback to `ffsubsync`; `manual`: open overlay picker |
| `alass_path`     | string path          | Path to `alass` executable. Empty or `null` falls back to `/usr/bin/alass`.                                 |
| `ffsubsync_path` | string path          | Path to `ffsubsync` executable. Empty or `null` falls back to `/usr/bin/ffsubsync`.                         |
| `ffmpeg_path`    | string path          | Path to `ffmpeg` (used for internal subtitle extraction). Empty or `null` falls back to `/usr/bin/ffmpeg`.  |

Default trigger is `Ctrl+Alt+S` via `shortcuts.triggerSubsync`.
Customize it there, or set it to `null` to disable.

### AniList

AniList integration is opt-in and disabled by default. Enable it to allow SubMiner to update watched episode progress after playback.

```json
{
  "anilist": {
    "enabled": true,
    "accessToken": ""
  }
}
```

| Option        | Values          | Description                                                             |
| ------------- | --------------- | ----------------------------------------------------------------------- |
| `enabled`     | `true`, `false` | Enable AniList post-watch progress updates (default: `false`)           |
| `accessToken` | string          | Optional explicit AniList access token override (default: empty string) |

When `enabled` is `true` and `accessToken` is empty, SubMiner opens an AniList setup helper window. Keep `enabled` as `false` to disable all AniList setup/update behavior.

Current post-watch behavior:

- SubMiner attempts an update near episode completion (`>=85%` watched and at least `10` minutes watched).
- Episode/title detection is `guessit`-first with fallback to SubMiner's filename parser.
- If `guessit` is unavailable, updates still work via fallback parsing but title matching can be less accurate.
- If embedded AniList auth UI fails to render, SubMiner opens the authorize URL in your default browser and shows fallback instructions in-app.
- Failed updates are retried with a persistent backoff queue in the background.

Setup flow details:

1. Set `anilist.enabled` to `true`.
2. Leave `anilist.accessToken` empty and restart SubMiner (or run `--anilist-setup`) to trigger setup.
3. Approve access in AniList.
4. Callback flow returns to SubMiner via `subminer://anilist-setup?...`, and SubMiner stores the token automatically.
   - Encryption backend: Linux defaults to `gnome-libsecret`.
     Override with `--password-store=<backend>` (for example `--password-store=basic_text`).

Token + detection notes:

- `anilist.accessToken` can be set directly in config; when blank, SubMiner uses the locally stored encrypted token from setup.
- Detection quality is best when `guessit` is installed and available on `PATH`.
- When `guessit` cannot parse or is missing, SubMiner falls back automatically to internal filename parsing.

AniList CLI commands:

- `--anilist-status`: print current AniList token resolution state and retry queue counters.
- `--anilist-logout`: clear stored AniList token from local persisted state.
- `--anilist-setup`: open AniList setup/auth flow helper window.
- `--anilist-retry-queue`: process one ready retry queue item immediately.

### Jellyfin

Jellyfin integration is optional and disabled by default. When enabled, SubMiner can authenticate, list libraries/items, and resolve direct/transcoded playback URLs for mpv launch.

```json
{
  "jellyfin": {
    "enabled": true,
    "serverUrl": "http://127.0.0.1:8096",
    "username": "",
    "remoteControlEnabled": true,
    "remoteControlAutoConnect": true,
    "autoAnnounce": false,
    "remoteControlDeviceName": "SubMiner",
    "defaultLibraryId": "",
    "directPlayPreferred": true,
    "directPlayContainers": ["mkv", "mp4", "webm", "mov", "flac", "mp3", "aac"],
    "transcodeVideoCodec": "h264"
  }
}
```

| Option                     | Values          | Description                                                                                                  |
| -------------------------- | --------------- | ------------------------------------------------------------------------------------------------------------ |
| `enabled`                  | `true`, `false` | Enable Jellyfin integration and CLI commands (default: `false`)                                              |
| `serverUrl`                | string (URL)    | Jellyfin server base URL                                                                                     |
| `username`                 | string          | Default username used by `--jellyfin-login`                                                                  |
| `deviceId`                 | string          | Client device id sent in auth headers (default: `subminer`)                                                  |
| `clientName`               | string          | Client name sent in auth headers (default: `SubMiner`)                                                       |
| `clientVersion`            | string          | Client version sent in auth headers (default: `0.1.0`)                                                       |
| `defaultLibraryId`         | string          | Default library id for `--jellyfin-items` when CLI value is omitted                                          |
| `remoteControlEnabled`     | `true`, `false` | Enable Jellyfin cast/remote-control session support                                                          |
| `remoteControlAutoConnect` | `true`, `false` | Auto-connect Jellyfin remote session on app startup (requires `jellyfin.enabled` and `remoteControlEnabled`) |
| `autoAnnounce`             | `true`, `false` | Auto-run cast-target visibility announce check on connect (default: `false`)                                 |
| `remoteControlDeviceName`  | string          | Device name shown in Jellyfin cast/device lists                                                              |
| `pullPictures`             | `true`, `false` | Enable poster/icon fetching for launcher Jellyfin pickers                                                    |
| `iconCacheDir`             | string          | Cache directory for launcher-fetched Jellyfin poster icons                                                   |
| `directPlayPreferred`      | `true`, `false` | Prefer direct stream URLs before transcoding                                                                 |
| `directPlayContainers`     | string[]        | Container allowlist for direct play decisions                                                                |
| `transcodeVideoCodec`      | string          | Preferred transcode video codec fallback (default: `h264`)                                                   |

Jellyfin auth session (`accessToken` + `userId`) is stored in local encrypted storage after login/setup.

- On Linux, token storage defaults to `gnome-libsecret` for `safeStorage`. Override with `--password-store=<backend>` on launcher/app invocations when needed.

Launcher subcommands:

- `subminer jellyfin` (or `subminer jf`) opens setup.
- `subminer jellyfin -l --server ... --username ... --password ...` logs in.
- `subminer jellyfin --logout` clears stored credentials.
- `subminer jellyfin -p` opens play picker.
- `subminer jellyfin -d` starts cast discovery mode.
- These launcher commands also accept `--password-store=<backend>` to override the launcher-app forwarded Electron switch.

See [Jellyfin Integration](/jellyfin-integration) for the full setup and cast-to-device guide.

Jellyfin remote auto-connect runs only when all three are `true`: `jellyfin.enabled`, `jellyfin.remoteControlEnabled`, and `jellyfin.remoteControlAutoConnect`.

### Discord Rich Presence

Discord Rich Presence is optional and disabled by default. When enabled, SubMiner publishes a polished activity card that reflects current media title, playback state, and session timer.

```json
{
  "discordPresence": {
    "enabled": true,
    "updateIntervalMs": 3000,
    "debounceMs": 750
  }
}
```

| Option             | Values          | Description                                                |
| ------------------ | --------------- | ---------------------------------------------------------- |
| `enabled`          | `true`, `false` | Enable Discord Rich Presence updates (default: `false`)    |
| `updateIntervalMs` | number          | Minimum interval between activity updates in milliseconds  |
| `debounceMs`       | number          | Debounce window for bursty playback events in milliseconds |

Setup steps:

1. Set `discordPresence.enabled` to `true`.
2. Restart SubMiner.

SubMiner uses a fixed official activity card style for all users:

- Details: current media title while playing (fallback: `Mining and crafting (Anki cards)` when idle/disconnected)
- State: `Playing mm:ss / mm:ss` or `Paused mm:ss / mm:ss` (fallback: `Idle`)
- Large image key/text: `subminer-logo` / `SubMiner`
- Small image key/text: `study` / `Sentence Mining`
- No activity button by default

Troubleshooting:

- If the card does not appear, verify Discord desktop app is running.
- If images do not render, confirm asset keys exactly match uploaded Discord asset names.
- If Discord is closed/not installed/disconnects, SubMiner continues running and quietly skips presence updates.

### Immersion Tracking

Enable or disable local immersion analytics stored in SQLite for mined subtitles and media sessions:

```json
{
  "immersionTracking": {
    "enabled": true,
    "dbPath": "",
    "batchSize": 25,
    "flushIntervalMs": 500,
    "queueCap": 1000,
    "payloadCapBytes": 256,
    "maintenanceIntervalMs": 86400000,
    "retention": {
      "eventsDays": 7,
      "telemetryDays": 30,
      "dailyRollupsDays": 365,
      "monthlyRollupsDays": 1825,
      "vacuumIntervalDays": 7
    }
  }
}
```

| Option                         | Values                        | Description                                                                                                 |
| ------------------------------ | ----------------------------- | ----------------------------------------------------------------------------------------------------------- |
| `enabled`                      | `true`, `false`               | Enable immersion tracking. Defaults to `true`.                                                              |
| `dbPath`                       | string                        | Optional SQLite database path. Leave empty to use default app-data path at `<config dir>/immersion.sqlite`. |
| `batchSize`                    | integer (`1`-`10000`)         | Buffered writes per transaction. Default `25`.                                                              |
| `flushIntervalMs`              | integer (`50`-`60000`)        | Maximum queue delay before flush. Default `500ms`.                                                          |
| `queueCap`                     | integer (`100`-`100000`)      | In-memory queue cap. Overflow drops oldest writes. Default `1000`.                                          |
| `payloadCapBytes`              | integer (`64`-`8192`)         | Event payload byte cap before truncation marker. Default `256`.                                             |
| `maintenanceIntervalMs`        | integer (`60000`-`604800000`) | Prune + rollup maintenance cadence. Default `86400000` (24h).                                               |
| `retention.eventsDays`         | integer (`1`-`3650`)          | Raw event retention window. Default `7` days.                                                               |
| `retention.telemetryDays`      | integer (`1`-`3650`)          | Telemetry retention window. Default `30` days.                                                              |
| `retention.dailyRollupsDays`   | integer (`1`-`36500`)         | Daily rollup retention window. Default `365` days.                                                          |
| `retention.monthlyRollupsDays` | integer (`1`-`36500`)         | Monthly rollup retention window. Default `1825` days (~5 years).                                            |
| `retention.vacuumIntervalDays` | integer (`1`-`3650`)          | Minimum spacing between `VACUUM` passes. Default `7` days.                                                  |

When `dbPath` is blank or omitted, SubMiner writes telemetry and session summaries to the default app-data location:

```text
<config directory>/immersion.sqlite
```

Set `dbPath` only if you want to relocate the database (for backup, syncing, or inspection workflows). The database is created when tracking starts for the first time.

See [Immersion Tracking Storage](/immersion-tracking) for schema details, query templates, retention/rollup behavior, and backend portability notes.

### YouTube Subtitle Generation

Set defaults used by the `subminer` launcher for YouTube subtitle extraction/transcription:

```json
{
  "youtubeSubgen": {
    "mode": "automatic",
    "whisperBin": "/path/to/whisper-cli",
    "whisperModel": "/path/to/ggml-model.bin",
    "primarySubLanguages": ["ja", "jpn"]
  }
}
```

| Option                | Values                                 | Description                                                                                                                                           |
| --------------------- | -------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------- |
| `mode`                | `"automatic"`, `"preprocess"`, `"off"` | `automatic`: play immediately and load generated subtitles in background; `preprocess`: generate before playback; `off`: disable launcher generation. |
| `whisperBin`          | string path                            | Path to `whisper.cpp` CLI binary used as fallback transcription engine.                                                                               |
| `whisperModel`        | string path                            | Path to whisper model used by fallback transcription.                                                                                                 |
| `primarySubLanguages` | string[]                               | Primary subtitle language priority for YouTube subtitle generation (default `["ja", "jpn"]`).                                                         |

YouTube language targets are derived from subtitle config:

- primary track: `youtubeSubgen.primarySubLanguages` (falls back to `["ja","jpn"]`)
- secondary track: `secondarySub.secondarySubLanguages` (falls back to English when empty)

Precedence for launcher defaults is: CLI flag > environment variable > `config.jsonc` > built-in default.
