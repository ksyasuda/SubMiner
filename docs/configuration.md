# Configuration

Settings are stored in `$XDG_CONFIG_HOME/SubMiner/config.jsonc` (or `~/.config/SubMiner/config.jsonc` when `XDG_CONFIG_HOME` is unset). For backward compatibility, SubMiner also reads existing configs from lowercase `subminer` directories.

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
- If the target file exists, SubMiner prompts to create a timestamped backup and overwrite.
- In non-interactive shells, use `--backup-overwrite` to explicitly back up and overwrite.
- `pnpm run generate:config-example` regenerates both repository `config.example.jsonc` and docs-served `/config.example.jsonc` from the same centralized defaults.
- `make generate-config` builds and runs the same default-config generator via local Electron.

Invalid config values are handled with warn-and-fallback behavior: SubMiner logs the bad key/value and continues with the default for that option.

### Configuration Options Overview

The configuration file includes several main sections:

- [**AnkiConnect**](#ankiconnect) - Automatic Anki card creation with media
- [**Auto-Start Overlay**](#auto-start-overlay) - Automatically show overlay on MPV connection
- [**Visible Overlay Subtitle Binding**](#visible-overlay-subtitle-binding) - Link visible overlay toggles to MPV subtitle visibility
- [**Auto Subtitle Sync**](#auto-subtitle-sync) - Sync current subtitle with `alass`/`ffsubsync`
- [**Invisible Overlay**](#invisible-overlay) - Startup visibility behavior for the invisible mining layer
- [**Jimaku**](#jimaku) - Jimaku API configuration and defaults
- [**Keybindings**](#keybindings) - MPV command shortcuts
- [**Runtime Option Palette**](#runtime-option-palette) - Live, session-only option toggles
- [**Secondary Subtitles**](#secondary-subtitles) - Dual subtitle track support
- [**Shortcuts Configuration**](#shortcuts-configuration) - Overlay keyboard shortcuts
- [**Subtitle Position**](#subtitle-position) - Overlay vertical positioning
- [**Subtitle Style**](#subtitle-style) - Appearance customization
- [**Texthooker**](#texthooker) - Control browser opening behavior
- [**WebSocket Server**](#websocket-server) - Built-in subtitle broadcasting server
- [**YouTube Subtitle Generation**](#youtube-subtitle-generation) - Launcher defaults for yt-dlp + local whisper fallback

### AnkiConnect

Enable automatic Anki card creation and updates with media generation:

```json
{
  "ankiConnect": {
    "enabled": true,
    "url": "http://127.0.0.1:8765",
    "pollingRate": 3000,
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
      "sentenceCardModel": "Japanese sentences",
      "sentenceCardSentenceField": "Sentence",
      "sentenceCardAudioField": "SentenceAudio"
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

| Option               | Values                                  | Description                                                                                                                               |
| -------------------- | --------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `enabled`            | `true`, `false`                         | Enable AnkiConnect integration (default: `false`)                                                                                         |
| `url`                | string (URL)                            | AnkiConnect API URL (default: `http://127.0.0.1:8765`)                                                                                    |
| `pollingRate`        | number (ms)                             | How often to check for new cards (default: `3000`)                                                                                        |
| `deck`               | string                                  | Anki deck to monitor for new cards                                                                                                        |
| `ankiConnect.nPlusOne.decks`     | array of strings                        | Decks used for N+1 known-word cache lookups. When omitted/empty, falls back to `ankiConnect.deck`.                                                   |
| `fields.audio`                  | string                                  | Card field for audio files (default: `ExpressionAudio`)                                                                                   |
| `fields.image`                  | string                                  | Card field for images (default: `Picture`)                                                                                                |
| `fields.sentence`               | string                                  | Card field for sentences (default: `Sentence`)                                                                                            |
| `fields.miscInfo`               | string                                  | Card field for metadata (default: `"MiscInfo"`, set to `null` to disable)                                                                 |
| `fields.translation`            | string                                  | Card field for sentence-card translation/back text (default: `SelectionText`)                                                            |
| `ai.enabled`            | `true`, `false`                         | Use AI translation for sentence cards. Also auto-attempted when secondary subtitle is missing.                                           |
| `ai.alwaysUseAiTranslation` | `true`, `false`                     | When `true`, always use AI translation even if secondary subtitles exist. When `false`, AI is used only when no secondary subtitle exists. |
| `ai.apiKey`             | string                                  | API key for your OpenAI-compatible endpoint (required for translation).                                                                    |
| `ai.model`              | string                                  | Model id for your OpenAI-compatible endpoint (default: `openai/gpt-4o-mini`).                                                             |
| `ai.baseUrl`            | string (URL)                            | OpenAI-compatible API base URL; accepts with or without `/v1`.                                                                            |
| `ai.targetLanguage`     | string                                  | Target language name used in translation prompt (default: `English`).                                                                     |
| `ai.systemPrompt`       | string                                  | System prompt used for translation (default returns translation text only).                                                               |
| `media.generateAudio`           | `true`, `false`                         | Generate audio clips from video (default: `true`)                                                                                         |
| `media.generateImage`           | `true`, `false`                         | Generate image/animation screenshots (default: `true`)                                                                                    |
| `media.imageType`               | `"static"`, `"avif"`                    | Image type: static screenshot or animated AVIF (default: `"static"`)                                                                      |
| `media.imageFormat`             | `"jpg"`, `"png"`, `"webp"`              | Image format (default: `"jpg"`)                                                                                                           |
| `media.imageQuality`            | number (1-100)                          | Image quality for JPG/WebP; PNG ignores this (default: `92`)                                                                              |
| `media.imageMaxWidth`           | number (px)                             | Optional max width for static screenshots. Unset keeps source width.                                                                       |
| `media.imageMaxHeight`          | number (px)                             | Optional max height for static screenshots. Unset keeps source height.                                                                      |
| `media.animatedFps`             | number (1-60)                           | FPS for animated AVIF (default: `10`)                                                                                                     |
| `media.animatedMaxWidth`        | number (px)                             | Max width for animated AVIF (default: `640`)                                                                                              |
| `media.animatedMaxHeight`       | number (px)                             | Optional max height for animated AVIF. Unset keeps source aspect-constrained height.                                                       |
| `media.animatedCrf`             | number (0-63)                           | CRF quality for AVIF; lower = higher quality (default: `35`)                                                                              |
| `media.audioPadding`            | number (seconds)                        | Padding around audio clip timing (default: `0.5`)                                                                                         |
| `media.fallbackDuration`        | number (seconds)                        | Default duration if timing unavailable (default: `3.0`)                                                                                   |
| `media.maxMediaDuration`        | number (seconds)                        | Max duration for generated media from multi-line copy (default: `30`, `0` to disable)                                                     |
| `behavior.overwriteAudio`       | `true`, `false`                         | Replace existing audio on updates; when `false`, new audio is appended/prepended per `behavior.mediaInsertMode` (default: `true`)        |
| `behavior.overwriteImage`       | `true`, `false`                         | Replace existing images on updates; when `false`, new images are appended/prepended per `behavior.mediaInsertMode` (default: `true`)      |
| `behavior.mediaInsertMode`      | `"append"`, `"prepend"`                 | Where to insert new media when overwrite is off (default: `"append"`)                                                                     |
| `behavior.highlightWord`        | `true`, `false`                         | Highlight the word in sentence context (default: `true`)                                                                                  |
| `ankiConnect.nPlusOne.highlightEnabled` | `true`, `false`                   | Enable fast local highlighting for words already known in Anki (default: `false`)                                                       |
| `ankiConnect.nPlusOne.nPlusOne`      | hex color string                        | Text color for the single target token to study when exactly one unknown candidate exists in a sentence (default: `"#c6a0f6"`). |
| `ankiConnect.nPlusOne.knownWord`      | hex color string                        | Legacy known-word color kept for backward compatibility (default: `"#a6da95"`).                                                     |
| `ankiConnect.nPlusOne.matchMode` | `"headword"`, `"surface"`      | Matching strategy for known-word highlighting (default: `"headword"`). `headword` uses token headwords; `surface` uses visible subtitle text. |
| `ankiConnect.nPlusOne.refreshMinutes`   | number                             | Minutes between known-word cache refreshes (default: `1440`)                                                                              |
| `ankiConnect.nPlusOne.decks`     | array of strings                        | Decks used by known-word cache refresh. Leave empty for compatibility with legacy `deck` scope.                                            |
| `behavior.notificationType`     | `"osd"`, `"system"`, `"both"`, `"none"` | Notification type on card update (default: `"osd"`)                                                                                       |
| `behavior.autoUpdateNewCards`   | `true`, `false`                         | Automatically update cards on creation (default: `true`)                                                                                  |
| `metadata.pattern`              | string                                  | Format pattern for metadata: `%f`=filename, `%F`=filename+ext, `%t`=time                                                                  |
| `isLapis`            | object                                  | Lapis/shared sentence-card config: `{ enabled, sentenceCardModel, sentenceCardSentenceField, sentenceCardAudioField }`                    |
| `isKiku`             | object                                  | Kiku-only config: `{ enabled, fieldGrouping, deleteDuplicateInAuto }` (shared sentence/audio/model settings are inherited from `isLapis`) |

**Kiku / Lapis Note Type Support:**

SubMiner supports the [Lapis](https://github.com/donkuri/lapis) and [Kiku](https://kiku.youyoumu.my.id/) note types. Both `isLapis.enabled` and `isKiku.enabled` can be true; Kiku takes precedence for grouping behavior, while sentence-card model/field settings come from `isLapis`.

When enabled, sentence cards automatically set `IsSentenceCard` to `"x"` and populate the `Expression` field. Audio cards set `IsAudioCard` to `"x"`.

Kiku extends Lapis with **field grouping** — when a duplicate card is detected (same Word/Expression), SubMiner merges the two cards' content into one using Kiku's `data-group-id` HTML structure, organizing each mining instance into separate pages within the note.

### N+1 Word Highlighting

When `ankiConnect.nPlusOne.highlightEnabled` is enabled, SubMiner builds a local cache of known words from Anki to highlight already learned tokens in subtitle rendering.

Known-word cache policy:

- Initial sync runs when the integration starts if the cache is missing or stale.
- `ankiConnect.nPlusOne.refreshMinutes` controls the minimum time between refreshes; between refreshes, cached words are reused without querying Anki.
- `ankiConnect.nPlusOne.nPlusOne` sets the color for the single target token when exactly one eligible unknown word exists.
- `ankiConnect.nPlusOne.knownWord` sets the legacy known-word highlight color for tokens already in Anki.
- `ankiConnect.nPlusOne.decks` accepts one or more decks. If empty, it uses the legacy single `ankiConnect.deck` value as scope.
- Cache state is persisted to `known-words-cache.json` under the app `userData` directory.
- The cache is automatically invalidated when the configured scope changes (for example, when deck changes).
- Cache lookups are in-memory. By default, token headwords are matched against cached `Expression` / `Word` values; set `ankiConnect.nPlusOne.matchMode` to `"surface"` for raw subtitle text matching.
- `ankiConnect.behavior.nPlusOne*` legacy keys (`nPlusOneHighlightEnabled`, `nPlusOneRefreshMinutes`, `nPlusOneMatchMode`) are deprecated and only kept for backward compatibility.
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

**Image Quality Notes:**

- `imageQuality` affects JPG and WebP only; PNG is lossless and ignores this setting
- JPG quality is mapped to FFmpeg's scale (2-31, lower = better)
- WebP quality uses FFmpeg's native 0-100 scale

### Manual Card Update Shortcuts

When `behavior.autoUpdateNewCards` is set to `false`, new cards are detected but not automatically updated. Use these keyboard shortcuts for manual control:

| Shortcut       | Action                                                                                                       |
| -------------- | ------------------------------------------------------------------------------------------------------------ |
| `Ctrl+C`       | Copy the current subtitle line to clipboard (preserves line breaks)                                          |
| `Ctrl+Shift+C` | Enter multi-copy mode. Press `1-9` to copy that many recent lines, or `Esc` to cancel. Timeout: 3 seconds    |
| `Ctrl+V`       | Update the last added Anki card using subtitles from clipboard                                               |
| `Ctrl+G`       | Trigger Kiku duplicate field grouping for the last added card (only when `behavior.autoUpdateNewCards` is `false`)    |
| `Ctrl+S`       | Create a sentence card from the current subtitle line                                                        |
| `Ctrl+Shift+S` | Enter multi-mine mode. Press `1-9` to create a sentence card from that many recent lines, or `Esc` to cancel |
| `Ctrl+Shift+V` | Cycle secondary subtitle display mode (hidden → visible → hover)                                             |
| `Ctrl+Shift+A` | Mark the last added Anki card as an audio card (sets IsAudioCard, SentenceAudio, Sentence, Picture)          |
| `Ctrl+Shift+O` | Open runtime options palette (session-only live toggles)                                                      |

**Multi-line copy workflow:**

1. Press `Ctrl+Shift+C`
2. Press a number key (`1-9`) within 3 seconds
3. The specified number of most recent subtitle lines are copied
4. Press `Ctrl+V` to update the last added card with the copied lines

These shortcuts are only active when the overlay window is visible and automatically disabled when hidden.

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

The mpv plugin controls startup per layer via `auto_start_visible_overlay` and `auto_start_invisible_overlay` in `subminer.conf` (`platform-default` for invisible means hidden on Linux, visible on macOS/Windows).

### Visible Overlay Subtitle Binding

Control whether toggling the visible overlay also toggles MPV subtitle visibility:

```json
{
  "bind_visible_overlay_to_mpv_sub_visibility": true
}
```

| Option                                        | Values          | Description |
| --------------------------------------------- | --------------- | ----------- |
| `bind_visible_overlay_to_mpv_sub_visibility` | `true`, `false` | When `true` (default), visible overlay hides MPV primary/secondary subtitles and restores them when hidden. When `false`, visible overlay toggles do not change MPV subtitle visibility. |

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

| Option           | Values               | Description |
| ---------------- | -------------------- | ----------- |
| `defaultMode`    | `"auto"`, `"manual"` | `auto`: try `alass` against secondary subtitle, then fallback to `ffsubsync`; `manual`: open overlay picker |
| `alass_path`     | string path          | Path to `alass` executable. Empty or `null` falls back to `/usr/bin/alass`. |
| `ffsubsync_path` | string path          | Path to `ffsubsync` executable. Empty or `null` falls back to `/usr/bin/ffsubsync`. |
| `ffmpeg_path`    | string path          | Path to `ffmpeg` (used for internal subtitle extraction). Empty or `null` falls back to `/usr/bin/ffmpeg`. |

Default trigger is `Ctrl+Alt+S` via `shortcuts.triggerSubsync`.
Customize it there, or set it to `null` to disable.

### Invisible Overlay

SubMiner includes a second subtitle mining layer that can be visually invisible while still interactive for Yomitan lookups.

- `invisibleOverlay.startupVisibility` values:
1. `"platform-default"`: hidden on Wayland, visible on Windows/macOS/other sessions.
2. `"visible"`: always shown on startup.
3. `"hidden"`: always hidden on startup.

Invisible subtitle positioning can be adjusted directly in the invisible layer:

- `Ctrl/Cmd+Shift+P` toggles position edit mode.
- Use arrow keys to move the invisible subtitle text.
- Press `Enter` or `Ctrl/Cmd+S` to save, or `Esc` to cancel.
- This edit-mode shortcut is fixed (not currently configurable in `shortcuts`/`keybindings`).

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

### Keybindings

Add a `keybindings` array to configure keyboard shortcuts that send commands to mpv:

See `config.example.jsonc` for detailed configuration options and more examples.

**Default keybindings:**

| Key               | Command                    | Description                           |
| ----------------- | -------------------------- | ------------------------------------- |
| `Space`           | `["cycle", "pause"]`       | Toggle pause                          |
| `ArrowRight`      | `["seek", 5]`              | Seek forward 5 seconds                |
| `ArrowLeft`       | `["seek", -5]`             | Seek backward 5 seconds               |
| `ArrowUp`         | `["seek", 60]`             | Seek forward 60 seconds               |
| `ArrowDown`       | `["seek", -60]`            | Seek backward 60 seconds              |
| `Shift+KeyH`      | `["sub-seek", -1]`         | Jump to previous subtitle             |
| `Shift+KeyL`      | `["sub-seek", 1]`          | Jump to next subtitle                 |
| `Ctrl+Shift+KeyH` | `["__replay-subtitle"]`    | Replay current subtitle, pause at end |
| `Ctrl+Shift+KeyL` | `["__play-next-subtitle"]` | Play next subtitle, pause at end      |
| `KeyQ`            | `["quit"]`                 | Quit mpv                              |
| `Ctrl+KeyW`       | `["quit"]`                 | Quit mpv                              |

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

**See `config.example.jsonc`** for more keybinding examples and configuration options.

### Runtime Option Palette

Use the runtime options palette to toggle settings live while SubMiner is running. These changes are session-only and reset on restart.

Current runtime options:

- `ankiConnect.behavior.autoUpdateNewCards` (`On` / `Off`)
- `ankiConnect.isKiku.fieldGrouping` (`auto` / `manual` / `disabled`)

Default shortcut: `Ctrl+Shift+O`

Palette controls:

- `Arrow Up/Down`: select option
- `Arrow Left/Right`: change selected value
- `Enter`: apply selected value
- `Esc`: close

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

### Shortcuts Configuration

Customize or disable the overlay keyboard shortcuts:

See `config.example.jsonc` for detailed configuration options.

```json
{
  "shortcuts": {
    "toggleVisibleOverlayGlobal": "Alt+Shift+O",
    "toggleInvisibleOverlayGlobal": "Alt+Shift+I",
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

| Option                        | Values           | Description                                                                                                                          |
| ----------------------------- | ---------------- | ------------------------------------------------------------------------------------------------------------------------------------ |
| `toggleVisibleOverlayGlobal`  | string \| `null` | Global accelerator for toggling visible subtitle overlay (default: `"Alt+Shift+O"`)                                               |
| `toggleInvisibleOverlayGlobal` | string \| `null` | Global accelerator for toggling invisible interactive overlay (default: `"Alt+Shift+I"`)                                         |
| `copySubtitle`                | string \| `null` | Accelerator for copying current subtitle (default: `"CommandOrControl+C"`)                                                           |
| `copySubtitleMultiple`        | string \| `null` | Accelerator for multi-copy mode (default: `"CommandOrControl+Shift+C"`)                                                              |
| `updateLastCardFromClipboard` | string \| `null` | Accelerator for updating card from clipboard (default: `"CommandOrControl+V"`)                                                       |
| `triggerFieldGrouping`        | string \| `null` | Accelerator for Kiku field grouping on last card (default: `"CommandOrControl+G"`; only active when `behavior.autoUpdateNewCards` is `false`) |
| `triggerSubsync`              | string \| `null` | Accelerator for running Subsync (default: `"Ctrl+Alt+S"`)                                                                           |
| `mineSentence`                | string \| `null` | Accelerator for creating sentence card from current subtitle (default: `"CommandOrControl+S"`)                                       |
| `mineSentenceMultiple`        | string \| `null` | Accelerator for multi-mine sentence card mode (default: `"CommandOrControl+Shift+S"`)                                                |
| `multiCopyTimeoutMs`          | number           | Timeout in ms for multi-copy/mine digit input (default: `3000`)                                                                      |
| `toggleSecondarySub`          | string \| `null` | Accelerator for cycling secondary subtitle mode (default: `"CommandOrControl+Shift+V"`)                                              |
| `markAudioCard`               | string \| `null` | Accelerator for marking last card as audio card (default: `"CommandOrControl+Shift+A"`)                                              |
| `openRuntimeOptions`          | string \| `null` | Opens runtime options palette for live session-only toggles (default: `"CommandOrControl+Shift+O"`)                                  |
| `openJimaku`                  | string \| `null` | Opens the Jimaku search modal (default: `"Ctrl+Shift+J"`)                                                                              |

**See `config.example.jsonc`** for the complete list of shortcut configuration options.

Set any shortcut to `null` to disable it.

### Subtitle Position

Set the initial vertical subtitle position (measured from the bottom of the screen):

```json
{
  "subtitlePosition": {
    "yPercent": 10
  }
}
```

| Option     | Values          | Description                                                        |
| ---------- | --------------- | ------------------------------------------------------------------ |
| `yPercent` | number (0 - 100) | Distance from the bottom as a percent of screen height (default: `10`) |

### Subtitle Style

Customize the appearance of primary and secondary subtitles:

See `config.example.jsonc` for detailed configuration options.

```json
{
  "subtitleStyle": {
    "fontFamily": "Noto Sans CJK JP Regular, Noto Sans CJK JP, Arial Unicode MS, Arial, sans-serif",
    "fontSize": 35,
    "fontColor": "#cad3f5",
    "fontWeight": "normal",
    "fontStyle": "normal",
    "backgroundColor": "rgba(54, 58, 79, 0.5)",
    "secondary": {
      "fontSize": 24,
      "fontColor": "#ffffff",
      "backgroundColor": "transparent"
    }
  }
}
```

| Option            | Values      | Description                                                                   |
| ----------------- | ----------- | ----------------------------------------------------------------------------- |
| `fontFamily`      | string      | CSS font-family value (default: `"Noto Sans CJK JP Regular, ..."`)            |
| `fontSize`        | number (px) | Font size in pixels (default: `35`)                                           |
| `fontColor`       | string      | Any CSS color value (default: `"#cad3f5"`)                                    |
| `fontWeight`      | string      | CSS font-weight, e.g. `"bold"`, `"normal"`, `"600"` (default: `"normal"`)     |
| `fontStyle`       | string      | `"normal"` or `"italic"` (default: `"normal"`)                                |
| `backgroundColor` | string      | Any CSS color, including `"transparent"` (default: `"rgba(54, 58, 79, 0.5)"`) |
| `enableJlpt`      | boolean     | Enable JLPT level underline styling (`false` by default)                        |
| `nPlusOneColor`   | string      | Existing n+1 highlight color (default: `#c6a0f6`)                            |
| `knownWordColor`  | string      | Existing known-word highlight color (default: `#a6da95`)                       |
| `jlptColors`      | object      | JLPT level underline colors object (`N1`..`N5`)                               |
| `secondary`       | object      | Override any of the above for secondary subtitles (optional)                  |

Secondary subtitle defaults: `fontSize: 24`, `fontColor: "#ffffff"`, `backgroundColor: "transparent"`. Any property not set in `secondary` falls back to the CSS defaults.

**See `config.example.jsonc`** for the complete list of subtitle style configuration options.

`jlptColors` keys are:

| Key  | Default   | Description                              |
| ---- | --------- | ---------------------------------------- |
| `N1` | `#ed8796` | JLPT N1 underline color                  |
| `N2` | `#f5a97f` | JLPT N2 underline color                  |
| `N3` | `#f9e2af` | JLPT N3 underline color                  |
| `N4` | `#a6e3a1` | JLPT N4 underline color                  |
| `N5` | `#8aadf4` | JLPT N5 underline color                  |

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

| Option         | Values                                  | Description |
| -------------- | --------------------------------------- | ----------- |
| `mode`         | `"automatic"`, `"preprocess"`, `"off"` | `automatic`: play immediately and load generated subtitles in background; `preprocess`: generate before playback; `off`: disable launcher generation. |
| `whisperBin`   | string path                             | Path to `whisper.cpp` CLI binary used as fallback transcription engine. |
| `whisperModel` | string path                             | Path to whisper model used by fallback transcription. |
| `primarySubLanguages` | string[]                          | Primary subtitle language priority for YouTube subtitle generation (default `["ja", "jpn"]`). |

YouTube language targets are derived from subtitle config:
- primary track: `youtubeSubgen.primarySubLanguages` (falls back to `["ja","jpn"]`)
- secondary track: `secondarySub.secondarySubLanguages` (falls back to English when empty)

Precedence for launcher defaults is: CLI flag > environment variable > `config.jsonc` > built-in default.
