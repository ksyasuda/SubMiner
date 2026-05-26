---
outline: [2, 3]
---

# Configuration

<script setup>
import { withBase } from 'vitepress';
</script>

SubMiner is configured through a single file (`config.jsonc`). Most settings are also editable from the in-app **Settings** window — you rarely need to edit the file by hand. This page is the full reference: it explains the Settings window, where the config file lives, and documents every option grouped by topic. New to SubMiner? The Quick Start below plus the [Settings window](#settings) cover everything most users need.

## Quick Start

For most users, start with this minimal configuration:

```json
{
  "ankiConnect": {
    "enabled": true,
    "deck": "YourDeckName",
    "knownWords": {
      "decks": {
        "YourDeckName": ["Word", "Word Reading", "Expression"]
      }
    },
    "fields": {
      "sentence": "Sentence",
      "audio": "Audio",
      "image": "Image"
    }
  }
}
```

`ankiConnect.deck` is still accepted for backward-compatible polling scope and legacy known-word fallback behavior. For known-word cache scope, prefer `ankiConnect.knownWords.decks` with deck-to-fields mapping.

Then customize as needed using the sections below.

## Settings

SubMiner includes a dedicated **Settings** window accessible from the tray menu, the app `--settings` flag, or launcher commands such as `subminer --settings` and `subminer settings`. It is the primary way to configure SubMiner — all changes are written directly to `config.jsonc`, so manual file editing is not required for most users.

The Settings window groups options by workflow instead of mirroring the raw config-file shape:

- Appearance
- Behavior
- Mining & Anki
- Playback & Sources
- Input
- Integrations
- Tracking & App
- Advanced

Each field still writes to its current `config.jsonc` path. For example, subtitle hover pause appears under **Behavior** / playback behavior, but saves to `subtitleStyle.autoPauseVideoOnHover`. Anki-aware fields can query AnkiConnect for deck names, note types, and field names, and keybinding fields use click-to-learn controls instead of raw text boxes.

The Settings window preserves existing JSONC comments, trailing commas, unrelated keys, and unsupported legacy options. Resetting a field removes the explicit config path so the built-in default applies.

Secret fields do not display stored values. They show whether a value is configured; entering a new value writes it, and reset clears the explicit path. Prefer command-based secret options such as `ai.apiKeyCommand` when available.

Saving validates the candidate config before writing. Live-reloadable changes are applied immediately; other changes return a restart-required banner in the window.

## Configuration File

The Settings window writes to `config.jsonc` directly, so most users do not need to edit the file by hand. The config file and the option reference below are provided for advanced use, scripting, or cases where you prefer editing config directly.

Settings are stored in `$XDG_CONFIG_HOME/SubMiner/config.jsonc` (or `~/.config/SubMiner/config.jsonc` when `XDG_CONFIG_HOME` is unset).
On Windows, the default path is `%APPDATA%\SubMiner\config.jsonc`.
When both files exist, SubMiner prefers `config.jsonc` over `config.json`.

See [config.example.jsonc](/config.example.jsonc) for a comprehensive example with all available options, default values, and detailed comments. Only include the options you want to customize in your config file.

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
- On Windows, generated configs default to `%APPDATA%\SubMiner\config.jsonc`.

Malformed config syntax (invalid JSON/JSONC) is startup-blocking: SubMiner shows a clear parse error with the config path and asks you to fix the file and restart.

For valid JSON/JSONC with invalid option values, SubMiner uses warn-and-fallback behavior: it logs the bad key/value and continues with the default for that option.

On macOS, these validation warnings also open a native dialog with full details (desktop notification banners can truncate long messages).

### Hot-Reload Behavior

SubMiner watches the active config file (`config.jsonc` or `config.json`) while running and applies supported updates automatically.

Hot-reloadable fields:

- `subtitleStyle`
- `subtitleSidebar`
- `keybindings`
- `shortcuts`
- `secondarySub.defaultMode`
- `stats.toggleKey`
- `stats.markWatchedKey`
- `logging.level`
- `youtube.primarySubLanguages`
- `jimaku.*`
- `subsync.*`
- `ankiConnect.ai.enabled`
- `ankiConnect.behavior.autoUpdateNewCards`
- `ankiConnect.knownWords.highlightEnabled`
- `ankiConnect.knownWords.refreshMinutes`
- `ankiConnect.knownWords.addMinedWordsImmediately`
- `ankiConnect.knownWords.matchMode`
- `ankiConnect.knownWords.decks`
- `ankiConnect.nPlusOne.enabled`
- `ankiConnect.nPlusOne.minSentenceWords`
- `ankiConnect.fields.word`
- `ankiConnect.fields.audio`
- `ankiConnect.fields.image`
- `ankiConnect.fields.sentence`
- `ankiConnect.fields.miscInfo`
- `ankiConnect.isLapis.sentenceCardModel`
- `ankiConnect.isKiku.fieldGrouping`

When these values change, SubMiner applies them live. Invalid config edits are rejected and the previous valid runtime config remains active.

Restart-required changes:

- Any other config sections still require restart.
- Shared top-level `ai` provider settings still require restart.
- AnkiConnect transport/proxy/media/deck/tag fields still require restart unless listed above.
- SubMiner shows an on-screen/system notification listing restart-required sections when they change.

### Configuration Options Overview

The configuration file includes several main sections:

**Core Settings**

- [**Logging**](#logging) - Runtime log level
- [**Auto-Start Overlay**](#auto-start-overlay) - Automatically show overlay on MPV connection
- [**Startup Warmups**](#startup-warmups) - Control what preloads on startup vs first-use defer
- [**WebSocket Server**](#websocket-server) - Built-in subtitle broadcasting server
- [**Annotation WebSocket**](#annotation-websocket) - Dedicated annotated subtitle payload stream
- [**Texthooker**](#texthooker) - Control browser opening behavior

**Subtitle Display**

- [**Subtitle Style**](#subtitle-style) - Appearance customization
- [**Subtitle Sidebar**](#subtitle-sidebar) - Parsed cue list sidebar modal
- [**Subtitle Position**](#subtitle-position) - Overlay vertical positioning
- [**Secondary Subtitles**](#secondary-subtitles) - Dual subtitle track support

**Keyboard & Controls**

- [**Keybindings**](#keybindings) - MPV command shortcuts
- [**Shortcuts Configuration**](#shortcuts-configuration) - Overlay keyboard shortcuts
- [**Controller Support**](#controller-support) - Gamepad support for keyboard-only mode
- [**Manual Card Update Shortcuts**](#manual-card-update-shortcuts) - Shortcuts for manual Anki card workflows
- [**Session Help Modal**](#session-help-modal) - In-overlay shortcut reference
- [**Runtime Option Palette**](#runtime-option-palette) - Live, session-only option toggles

**Anki Integration**

- [**Shared AI Provider**](#shared-ai-provider) - Canonical OpenAI-compatible provider config shared by Anki and YouTube subtitle fixing
- [**AnkiConnect**](#ankiconnect) - Automatic Anki card creation with media
- [**Kiku/Lapis Integration**](#kiku-lapis-integration) - Sentence cards and duplicate handling for Kiku/Lapis note types
- [**N+1 Word Highlighting**](#n1-word-highlighting) - Known-word cache and single-target highlighting
- [**Field Grouping Modes**](#field-grouping-modes) - Kiku/Lapis duplicate card merging

**External Integrations**

- [**Jimaku**](#jimaku) - Jimaku API configuration and defaults
- [**Subtitle Sync**](#subtitle-sync) - Sync current subtitle with `alass`/`ffsubsync`
- [**AniList**](#anilist) - Optional post-watch progress updates
- [**Yomitan**](#yomitan) - Reuse an external read-only Yomitan profile via `yomitan.externalProfilePath`
- [**Jellyfin**](#jellyfin) - Optional Jellyfin auth, library listing, and playback launch
- [**Discord Rich Presence**](#discord-rich-presence) - Optional Discord activity card updates
- [**Immersion Tracking**](#immersion-tracking) - Track subtitle sessions and mining activity in SQLite
- [**Stats Dashboard**](#stats-dashboard) - Local dashboard and overlay for immersion progress
- [**MPV Launcher**](#mpv-launcher) - mpv executable path, profile, and window launch mode
- [**YouTube Playback Settings**](#youtube-playback-settings) - Defaults for YouTube subtitle loading
- [**Updates**](#updates) - Automatic update checks, notifications, and prerelease testing

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

| Option  | Values                                   | Description                                               |
| ------- | ---------------------------------------- | --------------------------------------------------------- |
| `level` | `"debug"`, `"info"`, `"warn"`, `"error"` | Minimum log level for runtime logging (default: `"info"`) |

### Updates

Configure automatic update checks and update notifications:

```json
{
  "updates": {
    "enabled": true,
    "checkIntervalHours": 24,
    "notificationType": "system",
    "channel": "stable"
  }
}
```

| Option               | Values                                        | Description                                                                                         |
| -------------------- | --------------------------------------------- | --------------------------------------------------------------------------------------------------- |
| `enabled`            | `true`, `false`                               | Enable automatic background update checks. Manual tray and `subminer -u` checks are always allowed. |
| `checkIntervalHours` | number                                        | Minimum hours between automatic update checks. Default `24`.                                        |
| `notificationType`   | `"system"` \| `"osd"` \| `"both"` \| `"none"` | How SubMiner announces available updates. Default `"system"`.                                       |
| `channel`            | `"stable"` \| `"prerelease"`                  | Release channel used for update checks. Use `"prerelease"` to test beta/RC releases.                |

### Auto-Start Overlay

Control whether the overlay automatically becomes visible when it connects to mpv:

```json
{
  "auto_start_overlay": true
}
```

| Option               | Values          | Description                                           |
| -------------------- | --------------- | ----------------------------------------------------- |
| `auto_start_overlay` | `true`, `false` | Auto-show overlay on mpv connection (default: `true`) |

When you launch through the SubMiner app or the `subminer` wrapper, the launcher reads these settings from this config and injects them into the mpv plugin at runtime — there is no separate plugin config file to edit. `auto_start_overlay` controls whether the visible overlay shows on auto-start. Two related keys in the `mpv` block tune startup behavior: `mpv.autoStartSubMiner` starts the overlay automatically when a file loads, and `mpv.pauseUntilOverlayReady` pauses mpv on visible auto-start until SubMiner signals overlay/tokenization readiness.

On Windows, packaged plugin installs also rewrite the plugin socket path to `\\.\pipe\subminer-socket`.

### Startup Warmups

Control which startup warmups run in the background versus deferring to first real usage:

```json
{
  "startupWarmups": {
    "lowPowerMode": false,
    "mecab": true,
    "yomitanExtension": true,
    "subtitleDictionaries": true,
    "jellyfinRemoteSession": false
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

Defaults warm local tokenizer/dictionary work (`true` for `mecab`, `yomitanExtension`, and `subtitleDictionaries`) with `lowPowerMode: false`; Jellyfin remote session warmup is opt-in (`false` by default). Setting a warmup toggle to `false` defers that work until first usage.

### WebSocket Server

The overlay includes a built-in WebSocket server that broadcasts plain subtitle text to connected clients for external processing.

For endpoint details, payload examples, and client patterns, see [WebSocket / Texthooker API & Integration](/websocket-texthooker-api).

By default, the server is disabled. Set `enabled` to `true` to force it on, or `"auto"` to start it unless [mpv_websocket](https://github.com/kuroahna/mpv_websocket) is detected at `~/.config/mpv/mpv_websocket`.

See `config.example.jsonc` for detailed configuration options.

```json
{
  "websocket": {
    "enabled": false,
    "port": 6677
  }
}
```

| Option    | Values                    | Description                                         |
| --------- | ------------------------- | --------------------------------------------------- |
| `enabled` | `true`, `false`, `"auto"` | Built-in subtitle websocket mode (default: `false`) |
| `port`    | number                    | WebSocket server port (default: 6677)               |

### Annotation WebSocket

SubMiner also exposes a dedicated annotated websocket stream for the bundled texthooker UI and token-aware clients.

This stream includes subtitle text plus token metadata (N+1, known-word, frequency, JLPT, and character-name annotation context).

```json
{
  "annotationWebsocket": {
    "enabled": false,
    "port": 6678
  }
}
```

| Option    | Values          | Description                                                    |
| --------- | --------------- | -------------------------------------------------------------- |
| `enabled` | `true`, `false` | Toggle annotated websocket stream (independent of `websocket`) |
| `port`    | number          | Annotation websocket port (default: 6678)                      |

### Texthooker

Control whether texthooker starts automatically and whether it opens a browser:

See `config.example.jsonc` for detailed configuration options.

```json
{
  "texthooker": {
    "launchAtStartup": false,
    "openBrowser": false
  }
}
```

| Option            | Values          | Description                                                             |
| ----------------- | --------------- | ----------------------------------------------------------------------- |
| `launchAtStartup` | `true`, `false` | Start texthooker automatically with SubMiner startup (default: `false`) |
| `openBrowser`     | `true`, `false` | Open browser tab when texthooker starts (default: `false`)              |

## Subtitle Display

### Subtitle Style

Customize the appearance of primary and secondary subtitles:

See `config.example.jsonc` for detailed configuration options.

```json
{
  "subtitleStyle": {
    "fontColor": "#cad3f5",
    "backgroundColor": "transparent",
    "css": {
      "font-family": "Hiragino Sans, M PLUS 1, Source Han Sans JP, Noto Sans CJK JP",
      "font-size": "35px",
      "font-weight": "600",
      "line-height": "1.35",
      "letter-spacing": "-0.01em",
      "word-spacing": "0",
      "font-kerning": "normal",
      "text-rendering": "geometricPrecision",
      "text-shadow": "0 2px 6px rgba(0,0,0,0.9), 0 0 12px rgba(0,0,0,0.55)",
      "font-style": "normal",
      "backdrop-filter": "blur(6px)"
    },
    "secondary": {
      "fontColor": "#cad3f5",
      "backgroundColor": "transparent",
      "css": {
        "font-family": "Hiragino Sans, M PLUS 1, Source Han Sans JP, Noto Sans CJK JP",
        "font-size": "24px",
        "text-shadow": "0 2px 6px rgba(0,0,0,0.9), 0 0 12px rgba(0,0,0,0.55)"
      }
    }
  }
}
```

| Option                             | Values      | Description                                                                                                                          |
| ---------------------------------- | ----------- | ------------------------------------------------------------------------------------------------------------------------------------ |
| `fontFamily`                       | string      | CSS font-family value (default: `"Hiragino Sans, M PLUS 1, Source Han Sans JP, Noto Sans CJK JP"`)                                   |
| `fontSize`                         | number (px) | Font size in pixels (default: `35`)                                                                                                  |
| `fontColor`                        | string      | Any CSS color value (default: `"#cad3f5"`)                                                                                           |
| `css`                              | object      | CSS declarations applied to subtitles after normal style defaults; the settings window writes textbox edits here                     |
| `fontWeight`                       | string      | CSS font-weight, e.g. `"bold"`, `"normal"`, `"600"` (default: `"600"`)                                                               |
| `fontStyle`                        | string      | `"normal"` or `"italic"` (default: `"normal"`)                                                                                       |
| `backgroundColor`                  | string      | Any CSS color, including `"transparent"` (default: `"transparent"`)                                                                  |
| `enableJlpt`                       | boolean     | Enable JLPT level underline styling (`false` by default)                                                                             |
| `preserveLineBreaks`               | boolean     | Preserve line breaks in visible overlay subtitle rendering (`false` by default). Enable to mirror mpv line layout.                   |
| `autoPauseVideoOnHover`            | boolean     | Pause playback while mouse hovers subtitle text, then resume on leave (`true` by default).                                           |
| `autoPauseVideoOnYomitanPopup`     | boolean     | Pause playback while the Yomitan popup is open, then resume when the popup closes (`true` by default).                               |
| `hoverTokenColor`                  | string      | Hex color used for hovered subtitle token highlight in mpv (default: catppuccin mauve)                                               |
| `hoverTokenBackgroundColor`        | string      | CSS color used for hovered subtitle token background highlight (default: `"transparent"`); `hoverBackground` is accepted as an alias |
| `nameMatchEnabled`                 | boolean     | Enable subtitle token coloring for matches from the SubMiner character dictionary (`false` by default)                               |
| `nameMatchImagesEnabled`           | boolean     | Show small cached AniList character portraits beside matched character-name tokens (`false` by default)                              |
| `nameMatchColor`                   | string      | Hex color used for subtitle tokens matched from the SubMiner character dictionary (default: `#f5bde6`)                               |
| `knownWordColor`                   | string      | Hex color used for known-word subtitle highlights (default: `#a6da95`)                                                               |
| `nPlusOneColor`                    | string      | Hex color used for the single N+1 target subtitle highlight (default: `#c6a0f6`)                                                     |
| `frequencyDictionary.enabled`      | boolean     | Enable frequency highlighting from dictionary lookups (`false` by default)                                                           |
| `frequencyDictionary.sourcePath`   | string      | Path to a local frequency dictionary root. Leave empty or omit to use installed/default frequency-dictionary search paths.           |
| `frequencyDictionary.topX`         | number      | Only color tokens whose frequency rank is `<= topX` (`1000` by default)                                                              |
| `frequencyDictionary.mode`         | string      | `"single"` or `"banded"` (`"single"` by default)                                                                                     |
| `frequencyDictionary.matchMode`    | string      | `"headword"` or `"surface"` (`"headword"` by default)                                                                                |
| `frequencyDictionary.singleColor`  | string      | Color used for all highlighted tokens in single mode                                                                                 |
| `frequencyDictionary.bandedColors` | string[]    | Array of five hex colors used for ranked bands in banded mode                                                                        |
| `jlptColors`                       | object      | JLPT level underline colors object (`N1`..`N5`)                                                                                      |
| `secondary`                        | object      | Override any of the above for secondary subtitles (optional), including `secondary.css` declarations                                 |

The Settings window keeps subtitle color controls separate, then saves CSS textboxes to
`subtitleStyle.css`, `subtitleStyle.secondary.css`, and `subtitleSidebar.css`. The generated example
uses that same CSS declaration shape; existing top-level style keys such as `fontSize` and
`textShadow` remain supported for hand-written or older configs.

Frequency dictionary highlighting uses the same dictionary file format as JLPT bundle lookups (`term_meta_bank_*.json` under discovered dictionary directories). A token is highlighted when it has a positive integer `frequencyRank` (lower is more common) and the rank is within `topX`.

Lookup behavior:

- Set `frequencyDictionary.sourcePath` to a directory containing `term_meta_bank_*.json` for a fully custom source.
- If `sourcePath` is missing or empty, SubMiner searches default install/runtime locations for `frequency-dictionary` directories (for example app resources, user data paths, and current working directory).
- In both cases, only terms with a valid `frequencyRank` are used; everything else falls back to no highlighting.
- `frequencyDictionary.matchMode` controls which token text is used for frequency lookups: `headword` (dictionary form) or `surface` (visible subtitle text).
- Frequency highlighting skips tokens that look like non-lexical SFX/interjection noise (for example kana reduplication or short kana endings like `っ`), even when dictionary ranks exist.

In `single` mode all highlights use `singleColor`; in `banded` mode tokens map to five ascending color bands from most common to least common inside the topX window.

Character-name highlighting is separate from N+1 and frequency highlighting:

- `nameMatchEnabled` controls whether SubMiner includes character-dictionary name matches in subtitle token metadata and renderer styling.
- `nameMatchImagesEnabled` adds small circular portraits beside matched names using the AniList images already cached with character dictionary snapshots.
- `nameMatchColor` sets the highlight color for those matched character names.
- Matches come from the bundled SubMiner character dictionary, including AniList-synced merged dictionaries when enabled.

Secondary subtitle defaults: `fontFamily: "Hiragino Sans, M PLUS 1, Source Han Sans JP, Noto Sans CJK JP"`, `fontSize: 24`, `fontColor: "#cad3f5"`, `textShadow: "0 2px 6px rgba(0,0,0,0.9), 0 0 12px rgba(0,0,0,0.55)"`, `backgroundColor: "transparent"`, `fontWeight: "600"`. Any property not set in `secondary` falls back to the CSS defaults.

**See `config.example.jsonc`** for the complete list of subtitle style configuration options.

### Subtitle Sidebar

Configure the parsed-subtitle sidebar modal.

```json
{
  "subtitleSidebar": {
    "enabled": true,
    "autoOpen": false,
    "layout": "overlay",
    "toggleKey": "Backslash",
    "pauseVideoOnHover": true,
    "autoScroll": true,
    "fontFamily": "Hiragino Sans, M PLUS 1, Source Han Sans JP, Noto Sans CJK JP",
    "fontSize": 16
  }
}
```

| Option                      | Values    | Description                                                                                             |
| --------------------------- | --------- | ------------------------------------------------------------------------------------------------------- |
| `enabled`                   | boolean   | Enable subtitle sidebar support (`true` by default)                                                     |
| `autoOpen`                  | boolean   | Open sidebar automatically on overlay startup (`false` by default)                                      |
| `layout`                    | string    | `"overlay"` floats over mpv; `"embedded"` reserves right-side player space to mimic browser-like layout |
| `toggleKey`                 | string    | `KeyboardEvent.code` used to open/close the sidebar (default: `"Backslash"`)                            |
| `pauseVideoOnHover`         | boolean   | Pause playback while hovering the sidebar cue list (`true` by default)                                  |
| `autoScroll`                | boolean   | Keep the active cue in view while playback advances                                                     |
| `maxWidth`                  | number    | Maximum sidebar width in CSS pixels (default: `420`)                                                    |
| `opacity`                   | number    | Sidebar opacity between `0` and `1` (default: `0.95`)                                                   |
| `backgroundColor`           | string    | Sidebar shell background color                                                                          |
| `textColor`                 | hex color | Default cue text color                                                                                  |
| `fontFamily`                | string    | CSS `font-family` value applied to sidebar cue text                                                     |
| `fontSize`                  | number    | Base sidebar cue font size in CSS pixels (default: `16`)                                                |
| `timestampColor`            | hex color | Cue timestamp color                                                                                     |
| `activeLineColor`           | hex color | Active cue text color                                                                                   |
| `activeLineBackgroundColor` | string    | Active cue background color                                                                             |
| `hoverLineBackgroundColor`  | string    | Hovered cue background color                                                                            |

The sidebar is only available when the active subtitle source has been parsed into a cue list. Default colors use Catppuccin Macchiato with a semi-transparent shell so the panel stays readable without feeling like an opaque settings dialog.

`embedded` layout is intended to act like a split-pane view: it reserves player space with a right-side video margin and keeps interaction in both the player area and sidebar. If you see unexpected offset behavior in your environment, switch back to `overlay` to isolate sidebar placement.

For full details on layout modes, behavior, and the keyboard shortcut, see the [Subtitle Sidebar](/subtitle-sidebar) page.

`jlptColors` keys are:

| Key  | Default   | Description             |
| ---- | --------- | ----------------------- |
| `N1` | `#ed8796` | JLPT N1 underline color |
| `N2` | `#f5a97f` | JLPT N2 underline color |
| `N3` | `#f9e2af` | JLPT N3 underline color |
| `N4` | `#8bd5ca` | JLPT N4 underline color |
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

`secondarySub.secondarySubLanguages` also acts as the fallback secondary-language priority for managed startup subtitle selection on local playback and YouTube playback.

**Display modes:**

- **hidden** — Secondary subtitles not shown
- **visible** — Always visible at top of overlay
- **hover** — Only visible when hovering over the subtitle area (default)

**See `config.example.jsonc`** for additional secondary subtitle configuration options.

## Keyboard & Controls

### Keybindings

Add a `keybindings` array to configure keyboard shortcuts that send mpv commands or SubMiner session actions:

See `config.example.jsonc` for detailed configuration options and more examples.

**Default keybindings:**

| Key                  | Command                       | Description                             |
| -------------------- | ----------------------------- | --------------------------------------- |
| `Space`              | `["cycle", "pause"]`          | Toggle pause                            |
| `KeyF`               | `["cycle", "fullscreen"]`     | Toggle fullscreen                       |
| `KeyJ`               | `["cycle", "sid"]`            | Cycle primary subtitle track            |
| `Shift+KeyJ`         | `["cycle", "secondary-sid"]`  | Cycle secondary subtitle track          |
| `Ctrl+Alt+KeyP`      | `["__playlist-browser-open"]` | Open playlist browser                   |
| `Ctrl+Alt+KeyC`      | `["__youtube-picker-open"]`   | Open the manual YouTube subtitle picker |
| `ArrowRight`         | `["seek", 5]`                 | Seek forward 5 seconds                  |
| `ArrowLeft`          | `["seek", -5]`                | Seek backward 5 seconds                 |
| `ArrowUp`            | `["seek", 60]`                | Seek forward 60 seconds                 |
| `ArrowDown`          | `["seek", -60]`               | Seek backward 60 seconds                |
| `Shift+KeyH`         | `["sub-seek", -1]`            | Jump to previous subtitle               |
| `Shift+KeyL`         | `["sub-seek", 1]`             | Jump to next subtitle                   |
| `Shift+BracketLeft`  | `["__sub-delay-prev-line"]`   | Shift subtitle delay to previous cue    |
| `Shift+BracketRight` | `["__sub-delay-next-line"]`   | Shift subtitle delay to next cue        |
| `Ctrl+Shift+KeyH`    | `["__replay-subtitle"]`       | Replay current subtitle, pause at end   |
| `Ctrl+Shift+KeyL`    | `["__play-next-subtitle"]`    | Play next subtitle, pause at end        |
| `KeyQ`               | `["quit"]`                    | Quit mpv                                |
| `Ctrl+KeyW`          | `["quit"]`                    | Quit mpv                                |

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

**Special commands:** Commands prefixed with `__` are handled internally by the overlay rather than sent to mpv. `__playlist-browser-open` opens the split-pane playlist browser for the current file's parent directory and the live mpv queue. `__replay-subtitle` replays the current subtitle and pauses at its end. `__play-next-subtitle` seeks to the next subtitle, plays it, and pauses at its end. `__sub-delay-next-line` shifts subtitle delay so the active line aligns to the next cue start in the active subtitle source. `__sub-delay-prev-line` shifts subtitle delay so the active line aligns to the previous cue start. `__runtime-options-open` opens the runtime options palette. `__runtime-option-cycle:<id>[:next|prev]` cycles a runtime option value.

**Supported commands:** Any valid mpv JSON IPC command array (`["cycle", "pause"]`, `["seek", 5]`, `["script-binding", "..."]`, etc.)

For subtitle-position and subtitle-track proxy commands (`sub-pos`, `sid`, `secondary-sid`) and subtitle delay commands (`sub-delay`), SubMiner also shows an mpv OSD notification after the command runs.

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
    "openCharacterDictionaryManager": "CommandOrControl+D",
    "openRuntimeOptions": "CommandOrControl+Shift+O",
    "openSessionHelp": "CommandOrControl+Slash",
    "openControllerSelect": "Alt+C",
    "openControllerDebug": "Alt+Shift+C",
    "openJimaku": "Ctrl+Shift+J",
    "toggleSubtitleSidebar": "Backslash",
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
| `openCharacterDictionaryManager` | string \| `null` | Opens the loaded character dictionary manager (default: `"CommandOrControl+D"`)                                                             |
| `openRuntimeOptions`          | string \| `null` | Opens runtime options palette for live session-only toggles (default: `"CommandOrControl+Shift+O"`)                                           |
| `openSessionHelp`             | string \| `null` | Opens the in-overlay session help modal (default: `"CommandOrControl+Slash"`)                                                                 |
| `openControllerSelect`        | string \| `null` | Opens the controller config/remap modal (default: `"Alt+C"`)                                                                                  |
| `openControllerDebug`         | string \| `null` | Opens the controller debug modal (default: `"Alt+Shift+C"`)                                                                                   |
| `openJimaku`                  | string \| `null` | Opens the Jimaku search modal (default: `"Ctrl+Shift+J"`)                                                                                     |
| `toggleSubtitleSidebar`       | string \| `null` | Dispatches the subtitle sidebar toggle action (default: `"Backslash"`). `subtitleSidebar.toggleKey` remains the primary bare-key setting.     |

**See `config.example.jsonc`** for the complete list of shortcut configuration options.

Set any shortcut to `null` to disable it.

Feature-dependent shortcuts/keybindings only run when their related integration is enabled. For example, Anki/Kiku shortcuts require `ankiConnect.enabled` (and Kiku-specific behavior where applicable), and Jellyfin remote startup behavior requires Jellyfin to be enabled.

### Controller Support

SubMiner can read controllers through the Chrome Gamepad API and map them onto the existing keyboard-only overlay workflow.

Important behavior:

- Controller input is only active while keyboard-only mode is enabled.
- Keyboard-only mode continues to work normally without a controller.
- By default SubMiner uses the first connected controller.
- Fresh installs keep controller support disabled until you set `controller.enabled` to `true`.
- `Alt+C` opens the controller config modal by default, and you can remap that shortcut through `shortcuts.openControllerSelect`.
- The `Alt+C` config modal and `Alt+Shift+C` debug modal stay closed while controller support is disabled.
- Click the binding badge, edit pencil, or `Learn`, then press the next fresh button, trigger, or stick direction you want to bind for that overlay action.
- Click the reset button beside the edit pencil to restore one binding to the built-in default.
- Learned bindings are saved under `controller.profiles` for the selected controller id. Global `controller.bindings` remains the fallback for controllers without a profile.
- `Alt+Shift+C` opens the debug modal by default, and you can remap that shortcut through `shortcuts.openControllerDebug`.
- The debug modal shows raw axes/button values plus a ready-to-copy `buttonIndices` config block.
- `controller.buttonIndices` is a semantic reference/legacy mapping. Changing it does not rewrite the raw numeric descriptor values already stored under `controller.bindings`.
- Turning keyboard-only mode off clears the keyboard-only token highlight state.
- Closing the Yomitan popup clears the temporary native text-selection fill, but keeps controller token selection active.

```jsonc
{
  "controller": {
    "enabled": true,
    "preferredGamepadId": "",
    "preferredGamepadLabel": "",
    "smoothScroll": true,
    "scrollPixelsPerSecond": 900,
    "horizontalJumpPixels": 160,
    "stickDeadzone": 0.2,
    "triggerInputMode": "auto",
    "triggerDeadzone": 0.5,
    "repeatDelayMs": 320,
    "repeatIntervalMs": 120,
    "buttonIndices": {
      "select": 6,
      "buttonSouth": 0,
      "buttonEast": 1,
      "buttonWest": 2,
      "buttonNorth": 3,
      "leftShoulder": 4,
      "rightShoulder": 5,
      "leftStickPress": 9,
      "rightStickPress": 10,
      "leftTrigger": 6,
      "rightTrigger": 7,
    },
    "bindings": {
      "toggleLookup": { "kind": "button", "buttonIndex": 0 },
      "closeLookup": { "kind": "button", "buttonIndex": 1 },
      "toggleKeyboardOnlyMode": { "kind": "button", "buttonIndex": 3 },
      "mineCard": { "kind": "button", "buttonIndex": 2 },
      "quitMpv": { "kind": "button", "buttonIndex": 6 },
      "previousAudio": { "kind": "none" },
      "nextAudio": { "kind": "button", "buttonIndex": 5 },
      "playCurrentAudio": { "kind": "button", "buttonIndex": 4 },
      "toggleMpvPause": { "kind": "button", "buttonIndex": 9 },
      "leftStickHorizontal": { "kind": "axis", "axisIndex": 0, "dpadFallback": "horizontal" },
      "leftStickVertical": { "kind": "axis", "axisIndex": 1, "dpadFallback": "vertical" },
      "rightStickHorizontal": { "kind": "axis", "axisIndex": 3, "dpadFallback": "none" },
      "rightStickVertical": { "kind": "axis", "axisIndex": 4, "dpadFallback": "none" },
    },
    "profiles": {
      "Xbox Wireless Controller": {
        "label": "Xbox Wireless Controller",
        "bindings": {
          "toggleLookup": { "kind": "button", "buttonIndex": 0 },
          "mineCard": { "kind": "button", "buttonIndex": 2 },
        },
      },
    },
  },
}
```

Default logical mapping:

- Left stick up/down: scroll Yomitan popup
- Left stick left/right: move subtitle token selection
- Right stick up/down: page-jump through Yomitan popup
- Right stick left/right: unused by default
- `A`: toggle lookup
- `B`: close lookup
- `Y`: toggle keyboard-only mode
- `X`: mine card
- `Minus` / `Select`: quit mpv
- `L1`: play current Yomitan audio (falls back to the first available track)
- `R1`: move to the next available Yomitan audio track
- `L3`: toggle mpv pause
- `L2` / `R2`: unbound by default

Discrete bindings may use raw button indices or raw axis directions, and analog bindings use raw axis indices with optional D-pad fallback. The `Alt+C` learn flow writes those descriptors under `controller.profiles["<controller id>"]` for the selected controller. Manual edits are only needed when you want to script or copy exact mappings.

If you bind a discrete action to an axis manually, include `direction`:

```jsonc
{
  "controller": {
    "bindings": {
      "toggleLookup": { "kind": "axis", "axisIndex": 5, "direction": "positive" },
    },
  },
}
```

Treat `controller.buttonIndices` as reference-only unless you are still using legacy semantic bindings or copying values from the debug modal. Updating `controller.buttonIndices` alone does not rewrite the hardcoded raw numeric values already present in `controller.bindings` or `controller.profiles.*.bindings`. If you need a real remap, prefer the `Alt+C` learn flow so both the source and the descriptor shape stay correct.

If you choose to bind `L2` or `R2` manually, set `triggerInputMode` to `analog` and tune `triggerDeadzone` when your controller reports triggers as analog values instead of digital pressed/not-pressed buttons. `auto` accepts either style and remains the default.

If one controller reports non-standard raw button numbers, override `controller.profiles["<controller id>"].buttonIndices` using values from the `Alt+Shift+C` debug modal. Use global `controller.buttonIndices` only when the mapping should apply to every controller without a profile.

If you update this controller documentation or the generated controller examples, run `bun run docs:test` and `bun run docs:build` before merging.

Tune `scrollPixelsPerSecond`, `horizontalJumpPixels`, deadzones, repeat timing, and profile `buttonIndices` to match your controller. See [config.example.jsonc](/config.example.jsonc) for the full generated comments for every controller field.

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
| `Ctrl+D`       | Open loaded character dictionary manager                                                                           |
| `Ctrl+Shift+O` | Open runtime options palette (session-only live toggles)                                                           |
| `Ctrl/Cmd+A`   | Append clipboard video path to MPV playlist (fixed, not currently configurable)                                    |

**Multi-line copy workflow:**

1. Press `Ctrl+Shift+C`
2. Press a number key (`1-9`) within 3 seconds
3. The specified number of most recent subtitle lines are copied
4. Press `Ctrl+V` to update the last added card with the copied lines

These shortcuts are only active when the overlay window is visible and automatically disabled when hidden.

### Session Help Modal

The session help modal opens from the overlay with `Ctrl/Cmd+/` by default. The mpv plugin also exposes it through the `Y-H` chord (falling back to `Y-K` if needed). It shows the current session keybindings and color legend.

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
- `ankiConnect.knownWords.highlightEnabled` (`On` / `Off`)
- `subtitleStyle.enableJlpt` (`On` / `Off`)
- `subtitleStyle.frequencyDictionary.enabled` (`On` / `Off`)
- `ankiConnect.knownWords.matchMode` (`headword` / `surface`)
- `ankiConnect.isKiku.fieldGrouping` (`auto` / `manual` / `disabled`)

Annotation toggles (`nPlusOne`, `enableJlpt`, `frequencyDictionary.enabled`) only apply to new subtitle lines after the toggle. The currently displayed line is not re-tokenized in place.

Default shortcut: `Ctrl+Shift+O`

Palette controls:

- `Arrow Up/Down`: select option
- `Arrow Left/Right`: change selected value
- `Enter`: apply selected value
- `Esc`: close

## Anki Integration

### Shared AI Provider

This is the single, shared connection to an OpenAI-compatible LLM endpoint. Configure it **once** here at the top level, and SubMiner reuses it wherever AI is needed (today: Anki translation/enrichment). Per-feature toggles and prompt/model tweaks live in their own sections (for example `ankiConnect.ai`) and inherit this transport.

```json
{
  "ai": {
    "enabled": false,
    "apiKey": "",
    "apiKeyCommand": "",
    "model": "openai/gpt-4o-mini",
    "baseUrl": "https://openrouter.ai/api",
    "requestTimeoutMs": 15000
  }
}
```

| Option             | Values               | Description                                                                          |
| ------------------ | -------------------- | ------------------------------------------------------------------------------------ |
| `enabled`          | `true`, `false`      | Enable shared AI provider features (default: `false`)                                |
| `apiKey`           | string               | Static API key for the shared provider                                               |
| `apiKeyCommand`    | string               | Shell command used to resolve the API key (preferred over a plaintext `apiKey`)      |
| `model`            | string               | Default model identifier requested from the provider (default: `openai/gpt-4o-mini`) |
| `baseUrl`          | string (URL)         | OpenAI-compatible base URL (default: `https://openrouter.ai/api`)                    |
| `systemPrompt`     | string               | Default system prompt sent with requests (default: a translation-engine prompt)      |
| `requestTimeoutMs` | integer milliseconds | Shared request timeout (default: `15000`)                                            |

SubMiner uses the shared provider for:

- Anki translation/enrichment when `ankiConnect.ai.enabled` is `true`

### AnkiConnect

Enable automatic Anki card creation and updates with media generation:

```json
{
  "ankiConnect": {
    "enabled": true,
    "url": "http://127.0.0.1:8765",
    "pollingRate": 3000,
    "proxy": {
      "enabled": true,
      "host": "127.0.0.1",
      "port": 8766,
      "upstreamUrl": "http://127.0.0.1:8765"
    },
    "tags": ["SubMiner"],
    "deck": "Learning::Japanese",
    "fields": {
      "word": "Expression",
      "audio": "ExpressionAudio",
      "image": "Picture",
      "sentence": "Sentence",
      "miscInfo": "MiscInfo",
      "translation": "SelectionText"
    },
    "ai": {
      "enabled": false,
      "model": "openai/gpt-4o-mini",
      "systemPrompt": "Translate mined sentence text only."
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
      "audioPadding": 0,
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

| Option                                            | Values                                  | Description                                                                                                                                                                                         |
| ------------------------------------------------- | --------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `enabled`                                         | `true`, `false`                         | Enable AnkiConnect integration (default: `true`)                                                                                                                                                    |
| `url`                                             | string (URL)                            | AnkiConnect API URL (default: `http://127.0.0.1:8765`)                                                                                                                                              |
| `pollingRate`                                     | number (ms)                             | How often to check for new cards in polling mode (default: `3000`; ignored for direct proxy `addNote`/`addNotes` updates)                                                                           |
| `proxy.enabled`                                   | `true`, `false`                         | Enable local AnkiConnect-compatible proxy for push-based auto-enrichment (default: `true`)                                                                                                          |
| `proxy.host`                                      | string                                  | Bind host for local AnkiConnect proxy (default: `127.0.0.1`)                                                                                                                                        |
| `proxy.port`                                      | number                                  | Bind port for local AnkiConnect proxy (default: `8766`)                                                                                                                                             |
| `proxy.upstreamUrl`                               | string (URL)                            | Upstream AnkiConnect URL that proxy forwards to (default: `http://127.0.0.1:8765`)                                                                                                                  |
| `tags`                                            | array of strings                        | Tags automatically added to cards mined/updated by SubMiner (default: `['SubMiner']`; set `[]` to disable automatic tagging).                                                                       |
| `ankiConnect.deck`                                | string                                  | Legacy Anki polling/compatibility scope. Newer known-word cache scoping should use `ankiConnect.knownWords.decks`.                                                                                  |
| `ankiConnect.knownWords.decks`                    | object                                  | Deck→fields mapping for known-word cache queries (for example `{ "Kaishi 1.5k": ["Word", "Word Reading"] }`).                                                                                       |
| `fields.word`                                     | string                                  | Card field for mined word / expression text (default: `Expression`)                                                                                                                                 |
| `fields.audio`                                    | string                                  | Card field for audio files (default: `ExpressionAudio`)                                                                                                                                             |
| `fields.image`                                    | string                                  | Card field for images (default: `Picture`)                                                                                                                                                          |
| `fields.sentence`                                 | string                                  | Card field for sentences (default: `Sentence`)                                                                                                                                                      |
| `fields.miscInfo`                                 | string                                  | Card field for metadata (default: `"MiscInfo"`, set to `null` to disable)                                                                                                                           |
| `fields.translation`                              | string                                  | Card field for sentence-card translation/back text (default: `SelectionText`)                                                                                                                       |
| `ankiConnect.ai.enabled`                          | `true`, `false`                         | Use AI translation for sentence cards. Also auto-attempted when secondary subtitle is missing.                                                                                                      |
| `ankiConnect.ai.model`                            | string                                  | Optional model override for Anki AI translation/enrichment flows.                                                                                                                                   |
| `ankiConnect.ai.systemPrompt`                     | string                                  | Optional system prompt override for Anki AI translation/enrichment flows.                                                                                                                           |
| `media.generateAudio`                             | `true`, `false`                         | Generate audio clips from video (default: `true`)                                                                                                                                                   |
| `media.generateImage`                             | `true`, `false`                         | Generate image/animation screenshots (default: `true`)                                                                                                                                              |
| `media.imageType`                                 | `"static"`, `"avif"`                    | Image type: static screenshot or animated AVIF (default: `"static"`)                                                                                                                                |
| `media.imageFormat`                               | `"jpg"`, `"png"`, `"webp"`              | Image format (default: `"jpg"`)                                                                                                                                                                     |
| `media.imageQuality`                              | number (1-100)                          | Image quality for JPG/WebP; PNG ignores this (default: `92`)                                                                                                                                        |
| `media.imageMaxWidth`                             | number (px)                             | Optional max width for static screenshots. Unset keeps source width.                                                                                                                                |
| `media.imageMaxHeight`                            | number (px)                             | Optional max height for static screenshots. Unset keeps source height.                                                                                                                              |
| `media.animatedFps`                               | number (1-60)                           | FPS for animated AVIF (default: `10`)                                                                                                                                                               |
| `media.animatedMaxWidth`                          | number (px)                             | Max width for animated AVIF (default: `640`)                                                                                                                                                        |
| `media.animatedMaxHeight`                         | number (px)                             | Optional max height for animated AVIF. Unset keeps source aspect-constrained height.                                                                                                                |
| `media.animatedCrf`                               | number (0-63)                           | CRF quality for AVIF; lower = higher quality (default: `35`)                                                                                                                                        |
| `media.syncAnimatedImageToWordAudio`              | `true`, `false`                         | Whether animated AVIF includes an opening frame synced to sentence word-audio timing (default: `true`).                                                                                             |
| `media.audioPadding`                              | number (seconds)                        | Optional padding around audio clip timing (default: `0`). Animated AVIF clips freeze the first frame during leading audio padding.                                                                  |
| `media.fallbackDuration`                          | number (seconds)                        | Default duration if timing unavailable (default: `3.0`)                                                                                                                                             |
| `media.maxMediaDuration`                          | number (seconds)                        | Max duration for generated media from multi-line copy (default: `30`, `0` to disable)                                                                                                               |
| `behavior.overwriteAudio`                         | `true`, `false`                         | Replace existing audio on updates; when `false`, new audio is appended/prepended per `behavior.mediaInsertMode`; manual clipboard updates always replace generated sentence audio (default: `true`) |
| `behavior.overwriteImage`                         | `true`, `false`                         | Replace existing images on updates; when `false`, new images are appended/prepended per `behavior.mediaInsertMode` (default: `true`)                                                                |
| `behavior.mediaInsertMode`                        | `"append"`, `"prepend"`                 | Where to insert new media when overwrite is off (default: `"append"`)                                                                                                                               |
| `behavior.highlightWord`                          | `true`, `false`                         | Highlight the word in sentence context (default: `true`)                                                                                                                                            |
| `ankiConnect.knownWords.highlightEnabled`         | `true`, `false`                         | Enable fast local highlighting for words already known in Anki (default: `false`)                                                                                                                   |
| `ankiConnect.knownWords.addMinedWordsImmediately` | `true`, `false`                         | Add words from successful mines into the local known-word cache immediately (default: `true`)                                                                                                       |
| `ankiConnect.knownWords.matchMode`                | `"headword"`, `"surface"`               | Matching strategy for known-word highlighting (default: `"headword"`). `headword` uses token headwords; `surface` uses visible subtitle text.                                                       |
| `ankiConnect.knownWords.refreshMinutes`           | number                                  | Minutes between known-word cache refreshes (default: `1440`)                                                                                                                                        |
| `ankiConnect.knownWords.decks`                    | object                                  | Deck→fields mapping used for known-word cache query scope (e.g. `{ "Kaishi 1.5k": ["Word", "Word Reading"] }`).                                                                                     |
| `ankiConnect.nPlusOne.enabled`                    | `true`, `false`                         | Enable N+1 subtitle highlighting (highlights the one unknown word in a sentence). Independent from `knownWords.highlightEnabled`. Requires known-word cache data (default: `false`).                |
| `ankiConnect.nPlusOne.minSentenceWords`           | number                                  | Minimum number of words required in a sentence before single unknown-word N+1 highlighting can trigger (default: `3`).                                                                              |
| `behavior.notificationType`                       | `"osd"`, `"system"`, `"both"`, `"none"` | Notification type on card update (default: `"osd"`)                                                                                                                                                 |
| `behavior.autoUpdateNewCards`                     | `true`, `false`                         | Automatically update cards on creation (default: `true`)                                                                                                                                            |
| `metadata.pattern`                                | string                                  | Format pattern for metadata: `%f`=filename, `%F`=filename+ext, `%t`=time                                                                                                                            |
| `isLapis`                                         | object                                  | Lapis/shared sentence-card config: `{ enabled, sentenceCardModel }`. Sentence/audio field names are fixed to `Sentence` and `SentenceAudio`.                                                        |
| `isKiku`                                          | object                                  | Kiku-only config: `{ enabled, fieldGrouping, deleteDuplicateInAuto }` (shared sentence/audio/model settings are inherited from `isLapis`)                                                           |

`ankiConnect.ai` only controls feature-local enablement plus optional `model` / `systemPrompt` overrides.
API key resolution, base URL, and timeout live under the shared top-level [`ai`](#shared-ai-provider) config.

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

When `ankiConnect.knownWords.highlightEnabled` is enabled, SubMiner builds a local cache of known words from Anki to highlight already learned tokens in subtitle rendering.

Known-word cache policy:

- Initial sync runs when the integration starts if the cache is missing or stale.
- `ankiConnect.knownWords.refreshMinutes` controls the minimum time between refreshes; between refreshes, cached words are reused without querying Anki.
- `subtitleStyle.nPlusOneColor` sets the color for the single target token when exactly one eligible unknown word exists.
- `ankiConnect.nPlusOne.minSentenceWords` sets the minimum token count required in a sentence for N+1 highlighting (default: `3`).
- `subtitleStyle.knownWordColor` sets the known-word highlight color for tokens already in Anki.
- `ankiConnect.knownWords.decks` accepts an object keyed by deck name. If omitted or empty, it falls back to the legacy `ankiConnect.deck` single-deck scope.
- Cache state is persisted to `known-words-cache.json` under the app `userData` directory.
- The cache is automatically invalidated when the configured scope changes (for example, when deck changes).
- Cache lookups are in-memory. By default, token headwords are matched against cached `Expression` / `Word` values; set `ankiConnect.knownWords.matchMode` to `"surface"` for raw subtitle text matching.
- A known-word cache match always receives known-word highlighting, even when part-of-speech filters suppress N+1, frequency, or JLPT annotations for that token.
- Legacy moved keys under `ankiConnect.nPlusOne` (`highlightEnabled`, `refreshMinutes`, `matchMode`, `decks`, `knownWord`) and older `ankiConnect.behavior.nPlusOne*` keys are deprecated and only kept for backward compatibility.
- Legacy top-level `ankiConnect` migration keys (for example `audioField`, `generateAudio`, `imageType`) are compatibility-only, validated before mapping, and ignored with a warning when invalid.
- If AnkiConnect is unreachable, the cache remains in its previous state and an on-screen/system status message is shown.
- Known-word sync activity is logged at `INFO`/`DEBUG` level with the `anki` logger scope and includes scope, notes returned, and word counts.

To refresh roughly once per day, set:

```json
{
  "ankiConnect": {
    "knownWords": {
      "highlightEnabled": true,
      "refreshMinutes": 1440
    },
    "nPlusOne": {
      "minSentenceWords": 3
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
When the manual merge popup opens, SubMiner pauses playback and closes any open Yomitan popup first so the merge flow can take focus.

<video controls playsinline preload="metadata" :poster="withBase('/assets/kiku-integration-poster.jpg')" style="width: 100%; max-width: 960px;">
  <source :src="withBase('/assets/kiku-integration.webm')" type="video/webm" />
  Your browser does not support the video tag.
</video>

<a :href="withBase('/assets/kiku-integration.webm')" target="_blank" rel="noreferrer">Open demo in a new tab</a>

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

### Subtitle Sync

Sync the active subtitle track from the overlay picker using `alass` or `ffsubsync`. Both are **optional external tools** that must be installed separately and available on your `PATH` (or configured via the path options below).

- [`alass`](https://github.com/kaegi/alass) — fast, audio-independent sync using a secondary subtitle as reference
- [`ffsubsync`](https://github.com/smacke/ffsubsync) — audio-based sync using the video file as reference

```json
{
  "subsync": {
    "alass_path": "",
    "ffsubsync_path": "",
    "ffmpeg_path": "",
    "replace": true
  }
}
```

| Option           | Values          | Description                                                                                                               |
| ---------------- | --------------- | ------------------------------------------------------------------------------------------------------------------------- |
| `alass_path`     | string path     | Path to `alass` executable. Empty or `null` resolves from `PATH`. `alass` must be installed separately.                   |
| `ffsubsync_path` | string path     | Path to `ffsubsync` executable. Empty or `null` resolves from `PATH`. `ffsubsync` must be installed separately.           |
| `ffmpeg_path`    | string path     | Path to `ffmpeg` (used for internal subtitle extraction). Empty or `null` falls back to `/usr/bin/ffmpeg`.                |
| `replace`        | `true`, `false` | When `true` (default), overwrite the active subtitle file on successful sync. When `false`, write `<name>_retimed.<ext>`. |

Default trigger is `Ctrl+Alt+S` via `shortcuts.triggerSubsync`.
Customize it there, or set it to `null` to disable.

### AniList

AniList integration is opt-in and disabled by default. Enable it to allow SubMiner to update watched episode progress after playback.

```json
{
  "anilist": {
    "enabled": true,
    "accessToken": "",
    "characterDictionary": {
      "enabled": false,
      "refreshTtlHours": 168,
      "maxLoaded": 3,
      "evictionPolicy": "delete",
      "profileScope": "all",
      "collapsibleSections": {
        "description": false,
        "characterInformation": false,
        "voicedBy": false
      }
    }
  }
}
```

| Option                                                         | Values                  | Description                                                                                                   |
| -------------------------------------------------------------- | ----------------------- | ------------------------------------------------------------------------------------------------------------- |
| `enabled`                                                      | `true`, `false`         | Enable AniList post-watch progress updates (default: `false`)                                                 |
| `accessToken`                                                  | string                  | Optional explicit AniList access token override (default: empty string)                                       |
| `characterDictionary.enabled`                                  | `true`, `false`         | Enable automatic import/update of the merged SubMiner character dictionary for recent AniList media           |
| `characterDictionary.refreshTtlHours`                          | number                  | Legacy compatibility setting. Parsed and preserved, but merged dictionary retention is now usage-based        |
| `characterDictionary.maxLoaded`                                | number                  | Maximum number of most-recently-used AniList media snapshots included in the merged dictionary (default: `3`) |
| `characterDictionary.evictionPolicy`                           | `"delete"`, `"disable"` | Legacy compatibility setting. Parsed and preserved, but merged dictionary eviction is now usage-based         |
| `characterDictionary.collapsibleSections.description`          | `true`, `false`         | Open the Description section by default in generated dictionary entries                                       |
| `characterDictionary.collapsibleSections.characterInformation` | `true`, `false`         | Open the Character Information section by default in generated dictionary entries                             |
| `characterDictionary.collapsibleSections.voicedBy`             | `true`, `false`         | Open the Voiced by section by default in generated dictionary entries                                         |
| `characterDictionary.profileScope`                             | `"all"`, `"active"`     | Apply dictionary settings updates to all Yomitan profiles or only active profile                              |

When `enabled` is `true` and `accessToken` is empty, SubMiner opens an AniList setup helper window. Keep `enabled` as `false` to disable all AniList setup/update behavior.

Character dictionary sync behavior:

- Snapshot identity is still AniList **media ID**.
- Sync/import runs only for the currently watched media when media path/title changes.
- SubMiner keeps a most-recently-used list of synced AniList media snapshots and rebuilds one merged Yomitan dictionary from that active set.
- `maxLoaded` controls how many recent AniList media snapshots stay in the merged dictionary at once.
- The merged dictionary title stays stable as `SubMiner Character Dictionary`, so Yomitan sees one rotating dictionary instead of one dictionary per anime.

Current post-watch behavior:

- SubMiner attempts an update near episode completion using the shared default minimum watch ratio (`0.85`, or `>=85%`) from `src/shared/watch-threshold.ts`, and requires at least `10` minutes watched. The same ratio is also used by local episode watched state transitions.
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

### Yomitan

SubMiner normally uses its bundled Yomitan profile under the app config directory. If you want to reuse dictionaries and profile settings from another Electron app, point SubMiner at that app's Yomitan Electron profile in read-only mode.

For GameSentenceMiner on Linux, the default overlay profile path is typically `~/.config/gsm_overlay`.

```json
{
  "yomitan": {
    "externalProfilePath": "/home/you/.config/gsm_overlay"
  }
}
```

| Option                | Values      | Description                                                                                                                                                                                                    |
| --------------------- | ----------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `externalProfilePath` | string path | Optional absolute path, or a path beginning with `~` (expanded to your home directory), to another app's Yomitan Electron profile. SubMiner loads that profile read-only and reuses its dictionaries/settings. |

External-profile mode behavior:

- SubMiner uses the external profile's Yomitan extension/session instead of its local copy.
- SubMiner reads the external profile's currently active Yomitan profile selection and installed dictionaries.
- SubMiner does not open its own Yomitan settings window in this mode.
- SubMiner does not import, delete, or update dictionaries/settings in the external profile.
- SubMiner character-dictionary features are fully disabled in this mode, including auto-sync, manual generation, and subtitle-side character-dictionary annotations.
- First-run setup does not require any internal dictionaries while this mode is configured. If you later launch without `yomitan.externalProfilePath`, setup will require at least one internal Yomitan dictionary unless SubMiner already finds one.

### Jellyfin

Jellyfin integration is optional and disabled by default. When enabled, SubMiner can authenticate, list libraries/items, and resolve direct/transcoded playback URLs for mpv launch.

```json
{
  "jellyfin": {
    "enabled": true,
    "serverUrl": "http://127.0.0.1:8096",
    "recentServers": ["http://127.0.0.1:8096"],
    "username": "",
    "remoteControlEnabled": true,
    "remoteControlAutoConnect": true,
    "autoAnnounce": false,
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
| `recentServers`            | string[]        | Recent Jellyfin server URLs shown in setup; entries are trimmed, deduped, and capped at 5                    |
| `username`                 | string          | Default username used by `--jellyfin-login`                                                                  |
| `defaultLibraryId`         | string          | Default library id for `--jellyfin-items` when CLI value is omitted                                          |
| `remoteControlEnabled`     | `true`, `false` | Enable Jellyfin cast/remote-control session support                                                          |
| `remoteControlAutoConnect` | `true`, `false` | Auto-connect Jellyfin remote session on app startup (requires `jellyfin.enabled` and `remoteControlEnabled`) |
| `autoAnnounce`             | `true`, `false` | Auto-run cast-target visibility announce check on connect (default: `false`)                                 |
| `pullPictures`             | `true`, `false` | Enable poster/icon fetching for launcher Jellyfin pickers                                                    |
| `iconCacheDir`             | string          | Cache directory for launcher-fetched Jellyfin poster icons                                                   |
| `directPlayPreferred`      | `true`, `false` | Prefer direct stream URLs before transcoding                                                                 |
| `directPlayContainers`     | string[]        | Container allowlist for direct play decisions                                                                |
| `transcodeVideoCodec`      | string          | Preferred transcode video codec fallback (default: `h264`)                                                   |

Jellyfin auth session (`accessToken` + `userId`) is stored in local encrypted storage after login/setup. The legacy `jellyfin.accessToken`, `jellyfin.userId`, `jellyfin.clientName`, `jellyfin.deviceId`, `jellyfin.clientVersion`, and `jellyfin.remoteControlDeviceName` config keys are not resolver-backed settings in the current runtime. SubMiner reports the Jellyfin client as `SubMiner`, derives the Jellyfin device id and visible device name from the OS hostname, and owns the client version internally. The Settings window also hides low-level default library fields (`defaultLibraryId`) so normal setup stays focused on server, auth, playback, and remote-control behavior.

- On Linux, token storage defaults to `gnome-libsecret` for `safeStorage`. Override with `--password-store=<backend>` on launcher/app invocations when needed.

Launcher subcommands:

- `subminer jellyfin` (or `subminer jf`) opens setup.
- `subminer jellyfin -l --server ... --username ... --password ...` logs in.
- `subminer jellyfin --logout` clears stored credentials.
- `subminer jellyfin -p` opens play picker.
- `subminer jellyfin -d` starts cast discovery mode in background/tray mode.
- These launcher commands also accept `--password-store=<backend>` to override the launcher-app forwarded Electron switch.

See [Jellyfin Integration](/jellyfin-integration) for the full setup and cast-to-device guide.

Jellyfin remote auto-connect runs only when all three are `true`: `jellyfin.enabled`, `jellyfin.remoteControlEnabled`, and `jellyfin.remoteControlAutoConnect`.

Jellyfin playback auto-launched through SubMiner loads the mpv plugin the same way regular playback does, and shows the visible subtitle overlay automatically so `subtitleStyle` applies to subtitles selected from Jellyfin.

When Jellyfin is enabled with a server URL and SubMiner is running, the tray menu also shows a `Jellyfin Discovery` checkbox. It starts or stops discovery for the current runtime session only and does not write config. Starting discovery still requires a valid stored or environment-provided Jellyfin auth session.

### Discord Rich Presence

Discord Rich Presence is enabled by default. SubMiner publishes a polished activity card that reflects current media title, playback state, and session timer unless you turn it off.

```json
{
  "discordPresence": {
    "enabled": true,
    "presenceStyle": "default",
    "updateIntervalMs": 3000,
    "debounceMs": 750
  }
}
```

| Option             | Values                                           | Description                                                |
| ------------------ | ------------------------------------------------ | ---------------------------------------------------------- |
| `enabled`          | `true`, `false`                                  | Enable Discord Rich Presence updates (default: `true`)     |
| `presenceStyle`    | `"default"`, `"meme"`, `"japanese"`, `"minimal"` | Card text preset (default: `"default"`)                    |
| `updateIntervalMs` | number                                           | Minimum interval between activity updates in milliseconds  |
| `debounceMs`       | number                                           | Debounce window for bursty playback events in milliseconds |

Setup steps:

1. Leave `discordPresence.enabled` as `true` or set it explicitly if you previously disabled it.
2. Optionally set `discordPresence.presenceStyle` to choose a card text preset.
3. Restart SubMiner.

#### Presence style presets

While playing media, the **Details** line always shows the current media title and **State** shows `Playing mm:ss / mm:ss` or `Paused mm:ss / mm:ss`. The preset controls what appears when idle and the tooltip text on images.

| Preset        | Idle details                       | Small image text   | Vibe                                    |
| ------------- | ---------------------------------- | ------------------ | --------------------------------------- |
| **`default`** | `Sentence Mining`                  | `日本語学習中`     | Clean, bilingual flair                  |
| `meme`        | `Mining and crafting (Anki cards)` | `Sentence Mining`  | Minecraft-inspired joke                 |
| `japanese`    | `文の採掘中`                       | `イマージョン学習` | Fully Japanese                          |
| `minimal`     | `SubMiner`                         | _(none)_           | Bare essentials, no small image overlay |

All presets use the `subminer-logo` large image with `SubMiner` tooltip. No activity button is shown by default.

Troubleshooting:

- If the card does not appear, verify Discord desktop app is running.
- If images do not render, confirm asset keys exactly match uploaded Discord asset names.
- If Discord is closed/not installed/disconnects, SubMiner continues running and quietly skips presence updates.

### Immersion Tracking

Enable or disable local immersion analytics stored in SQLite for mined subtitles and media sessions. This data also powers the stats dashboard:

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
    "retentionMode": "preset",
    "retentionPreset": "balanced",
    "retention": {
      "eventsDays": 0,
      "telemetryDays": 0,
      "sessionsDays": 0,
      "dailyRollupsDays": 0,
      "monthlyRollupsDays": 0,
      "vacuumIntervalDays": 0
    },
    "lifetimeSummaries": {
      "global": true,
      "anime": true,
      "media": true
    }
  }
}
```

| Option                         | Values                              | Description                                                                                                 |
| ------------------------------ | ----------------------------------- | ----------------------------------------------------------------------------------------------------------- |
| `enabled`                      | `true`, `false`                     | Enable immersion tracking. Defaults to `true`.                                                              |
| `dbPath`                       | string                              | Optional SQLite database path. Leave empty to use default app-data path at `<config dir>/immersion.sqlite`. |
| `batchSize`                    | integer (`1`-`10000`)               | Buffered writes per transaction. Default `25`.                                                              |
| `flushIntervalMs`              | integer (`50`-`60000`)              | Maximum queue delay before flush. Default `500ms`.                                                          |
| `queueCap`                     | integer (`100`-`100000`)            | In-memory queue cap. Overflow drops oldest writes. Default `1000`.                                          |
| `payloadCapBytes`              | integer (`64`-`8192`)               | Event payload byte cap before truncation marker. Default `256`.                                             |
| `maintenanceIntervalMs`        | integer (`60000`-`604800000`)       | Prune + rollup maintenance cadence. Default `86400000` (24h).                                               |
| `retentionMode`                | `preset`,`advanced`                 | Retention mode. `preset` applies `retentionPreset`, `advanced` uses explicit values only. Default `preset`. |
| `retentionPreset`              | `minimal`,`balanced`,`deep-history` | Retention preset used when `retentionMode = "preset"`. Default `balanced`.                                  |
| `retention.eventsDays`         | integer (`0`-`3650`)                | Raw event retention window in days. Default `0` (keep all).                                                 |
| `retention.telemetryDays`      | integer (`0`-`3650`)                | Telemetry retention window in days. Default `0` (keep all).                                                 |
| `retention.sessionsDays`       | integer (`0`-`3650`)                | Session retention window in days. Default `0` (keep all).                                                   |
| `retention.dailyRollupsDays`   | integer (`0`-`36500`)               | Daily rollup retention window. Default `0` (keep all).                                                      |
| `retention.monthlyRollupsDays` | integer (`0`-`36500`)               | Monthly rollup retention window. Default `0` (keep all).                                                    |
| `retention.vacuumIntervalDays` | integer (`0`-`3650`)                | Minimum spacing between `VACUUM` passes. `0` disables vacuum. Default `0` (disabled).                       |

You can also disable immersion tracking for a single session using:

```bash
SUBMINER_DISABLE_IMMERSION_TRACKING=1 subminer
```

When this is set, SubMiner skips immersion-tracker startup and does not initialize or read the immersion SQLite database for that session.

Default behavior keeps raw events, telemetry, sessions, and rollups forever while still maintaining lifetime summary tables and daily/monthly rollups for faster reads. If you later want bounded retention, switch `retentionMode` or set explicit `retention.*` values.

When `dbPath` is blank or omitted, SubMiner writes telemetry and session summaries to the default app-data location:

```text
<config directory>/immersion.sqlite
```

Set `dbPath` only if you want to relocate the database (for backup, syncing, or inspection workflows). The database is created when tracking starts for the first time.

See [Immersion Tracking Storage](/immersion-tracking) for schema details, query templates, dashboard access, retention/rollup behavior, backend portability notes, and the dedicated SQLite verification command.

### Stats Dashboard

Configure the local stats UI served from SubMiner and the in-app stats overlay toggle:

```json
{
  "stats": {
    "toggleKey": "Backquote",
    "serverPort": 6969,
    "autoStartServer": true,
    "autoOpenBrowser": false
  }
}
```

| Option            | Values            | Description                                                                                                          |
| ----------------- | ----------------- | -------------------------------------------------------------------------------------------------------------------- |
| `toggleKey`       | Electron key code | Overlay-local key code used to toggle the stats overlay. Default `Backquote`.                                        |
| `serverPort`      | integer           | Localhost port for the browser stats UI. Default `6969`.                                                             |
| `autoStartServer` | `true`, `false`   | Start the local stats HTTP server automatically once immersion tracking is active. Default `true`.                   |
| `autoOpenBrowser` | `true`, `false`   | When `subminer stats` starts the server on demand, also open the dashboard in your default browser. Default `false`. |

Usage notes:

- The browser UI is served at `http://127.0.0.1:<serverPort>`.
- The overlay toggle is local to the focused visible overlay window; it is not registered as a global OS shortcut.
- The dashboard reads from the same immersion-tracking database, so keep `immersionTracking.enabled` on if you want data to appear.
- The UI includes Overview, Library, Trends, Vocabulary, and Sessions tabs.

### MPV Launcher

Configure the mpv executable, profile, and window state for SubMiner-managed mpv launches (launcher playback, Windows `--launch-mpv`, and Jellyfin idle mpv startup):

```json
{
  "mpv": {
    "executablePath": "",
    "profile": "",
    "launchMode": "normal"
  }
}
```

| Option           | Values                                        | Description                                                                                                                         |
| ---------------- | --------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------- |
| `executablePath` | string                                        | Absolute path to `mpv.exe` for Windows launch flows. Leave empty to auto-discover from `SUBMINER_MPV_PATH` or `PATH` (default `""`) |
| `profile`        | string                                        | mpv profile name passed as `--profile=<name>`. Leave empty to pass no profile (default `""`)                                        |
| `launchMode`     | `"normal"` \| `"maximized"` \| `"fullscreen"` | Window state when SubMiner spawns mpv (default `"normal"`)                                                                          |

If `mpv.profile` is configured and the launcher also receives `--profile`, SubMiner passes both as a comma-separated mpv profile list.

Launch mode behavior:

- **`normal`** — mpv opens at its default window size with no extra flags.
- **`maximized`** — mpv starts maximized via `--window-maximized=yes`, keeping taskbar access.
- **`fullscreen`** — mpv starts in true fullscreen via `--fullscreen`.

### YouTube Playback Settings

Set defaults used by managed subtitle auto-selection and the `subminer` launcher YouTube flow:

```json
{
  "youtube": {
    "primarySubLanguages": ["ja", "jpn"]
  }
}
```

| Option                | Values   | Description                                                                                      |
| --------------------- | -------- | ------------------------------------------------------------------------------------------------ |
| `primarySubLanguages` | string[] | Primary subtitle language priority for managed subtitle auto-selection (default `["ja", "jpn"]`) |

Current launcher behavior:

- For YouTube URLs, SubMiner probes subtitle tracks with yt-dlp after mpv bootstrap and binds auto-selected tracks before normal playback resumes.
- If YouTube/mpv already exposes an authoritative matching subtitle track, SubMiner reuses it; otherwise it downloads and injects only the missing side.
- SubMiner loads the primary subtitle plus a best-effort secondary subtitle.
- Playback waits only for primary subtitle readiness; secondary failures do not block playback.
- English secondary subtitles are selected from `secondarySub.secondarySubLanguages` when primary language matches are unavailable.
- Native mpv secondary subtitle rendering stays hidden during this flow so the SubMiner overlay remains the visible secondary subtitle surface.
- If primary subtitle loading fails, use `Ctrl+Alt+C` to open the subtitle modal and pick a track.

Language targets are derived from subtitle config:

- primary track: `youtube.primarySubLanguages` (falls back to `["ja","jpn"]`)
- secondary track: `secondarySub.secondarySubLanguages` (falls back to English when empty)
- Local playback uses the same priorities after mpv reports subtitle track metadata, so sidecar/internal mixed sets can override an incorrect initial `sid=auto` pick.
- Tracks are resolved and loaded before mpv starts; the older launcher mode switch has been removed.

Precedence for launcher defaults is: CLI flag > environment variable > `config.jsonc` > built-in default.
