---
outline: [2, 3]
---

# Configuration

<script setup>
import { withBase } from 'vitepress';
</script>

All SubMiner settings live in one file, `config.jsonc`. Most of them are also editable in the Settings window, so you rarely need to edit the file by hand. This page lists every config block with its keys and defaults.

## Config file {#configuration-file}

| Platform     | Path                                                                                  |
| ------------ | ------------------------------------------------------------------------------------- |
| Linux, macOS | `$XDG_CONFIG_HOME/SubMiner/config.jsonc` (`~/.config/SubMiner/config.jsonc` if unset) |
| Windows      | `%APPDATA%\SubMiner\config.jsonc`                                                     |

The file is JSONC, so comments and trailing commas are allowed. If both `config.jsonc` and `config.json` exist, SubMiner uses `config.jsonc`. Only add the keys you want to change. Everything else uses the built-in default.

The [generated example config](/config.example.jsonc) lists every option with its default and a comment. Defaults in the tables below come from that file.

::: warning mpv.socketPath differs on Windows
The example shows `"socketPath": "/tmp/subminer-socket"`. On Windows the default is `\\.\pipe\subminer-socket`. Leave `mpv.socketPath` out of your config unless you need a custom path, and SubMiner picks the right one.
:::

To write a fresh default config:

```bash
SubMiner.AppImage --generate-config
SubMiner.AppImage --generate-config --config-path /tmp/subminer.jsonc
SubMiner.AppImage --generate-config --backup-overwrite
```

If the target file exists, SubMiner asks before backing it up and overwriting it. In non-interactive shells, pass `--backup-overwrite`.

A syntax error in the file stops startup with a message that names the file. A valid file with a bad value logs a warning and uses the default for that key. On macOS, these warnings also open a dialog.

## Settings window {#settings}

Open it from the tray menu, with `subminer settings`, or with the app's `--settings` flag. Options are grouped by task (Appearance, Behavior, Mining & Anki, Input, Integrations, Tracking & App, Advanced) rather than by config block, but each field saves to its normal `config.jsonc` path.

- Saving keeps your comments, trailing commas, and unrelated keys. Resetting a field removes its key so the default applies.
- Each field is tagged **Live** or **Restart**. After saving, a banner lists any sections that need a restart.
- Anki fields can fetch deck, note type, and field names from AnkiConnect.
- Secret fields never show the stored value, only whether one is set. Prefer the `*Command` variants (such as `jimaku.apiKeyCommand`) to keep keys out of the file.

## Hot-reload {#hot-reload-behavior}

SubMiner watches the config file while running. When it changes, live settings apply immediately and SubMiner shows a notification listing any changed sections that need a restart. If the new file is invalid, the previous config stays active.

These apply live:

- `subtitleStyle`, `subtitleSidebar`, `subtitleSelection`, `keybindings`, `shortcuts`
- `logging.level`, `logging.rotation`, `logging.files`
- `secondarySub.defaultMode`, `youtube.primarySubLanguages`
- `mpv.aniskipEnabled`, `mpv.aniskipButtonKey`, `stats.toggleKey`, `stats.markWatchedKey`
- `ankiConnect.deck`, `ankiConnect.fields.*`, `ankiConnect.behavior.autoUpdateNewCards`
- `ankiConnect.media.normalizeAudio`, `media.mirrorMpvVolume`, `media.reviewTiming`
- `ankiConnect.knownWords` (`highlightEnabled`, `refreshMinutes`, `addMinedWordsImmediately`, `matchMode`, `decks`) and `ankiConnect.nPlusOne.*`
- `ankiConnect.isLapis.sentenceCardModel`, `isKiku.fieldGrouping`, `isSenren.fieldGrouping`, `lapisKiku.wordCardKind`

These are read at the start of the next operation, so changes take effect on the next request or run: `jimaku`, `tmdb`, `subsync`, `subtitleGeneration`, `notifications`.

Everything else needs a restart.

## Core settings

### Logging

Log files are named by date (`app-YYYY-MM-DD.log`, `launcher-...`, `mpv-...`). Log export writes a sanitized copy and leaves the originals alone.

| Key                      | Default  | What it does                                            |
| ------------------------ | -------- | ------------------------------------------------------- |
| `logging.level`          | `"warn"` | Minimum level: `debug`, `info`, `warn`, `error`         |
| `logging.rotation`       | `7`      | Days of logs to keep                                    |
| `logging.files.app`      | `true`   | Write app logs                                          |
| `logging.files.launcher` | `true`   | Write launcher logs                                     |
| `logging.files.mpv`      | `false`  | Write mpv logs. Turn on temporarily to debug mpv/plugin |

### Updates

Manual checks from the tray or `subminer -u` always work, even with automatic checks off. Overlay update notifications include an **Update** button.

| Key                          | Default     | What it does                                              |
| ---------------------------- | ----------- | --------------------------------------------------------- |
| `updates.enabled`            | `true`      | Check for updates in the background                       |
| `updates.checkIntervalHours` | `24`        | Minimum hours between automatic checks                    |
| `updates.notificationType`   | `"overlay"` | `overlay`, `system`, `both` (overlay + system), or `none` |
| `updates.channel`            | `"stable"`  | `stable` or `prerelease` (betas and release candidates)   |

### Notifications

Overlay notifications are also kept in a session-only history panel. Toggle it with `shortcuts.toggleNotificationHistory`. The panel opens from the same side as the notification cards.

| Key                             | Default       | What it does                                               |
| ------------------------------- | ------------- | ---------------------------------------------------------- |
| `notifications.overlayPosition` | `"top-right"` | Where overlay cards appear: `top-left`, `top`, `top-right` |

Mining and startup status notifications use `ankiConnect.behavior.notificationType` (see [AnkiConnect](#ankiconnect)).

### Auto-start overlay

When mpv is started by SubMiner or the `subminer` launcher, the launcher passes these settings to the bundled mpv plugin. There is no separate plugin config file. `mpv.autoStartSubMiner` and `mpv.pauseUntilOverlayReady` (see [MPV launcher](#mpv-launcher)) control the background start and the initial pause.

| Key                  | Default | What it does                                                 |
| -------------------- | ------- | ------------------------------------------------------------ |
| `auto_start_overlay` | `true`  | Show the visible overlay when the mpv plugin starts SubMiner |

### Startup warmups

Warmups load components in the background at startup. Turn one off to load it on first use instead.

| Key                                    | Default | What it does                                                                  |
| -------------------------------------- | ------- | ----------------------------------------------------------------------------- |
| `startupWarmups.lowPowerMode`          | `false` | Defer every warmup except the Yomitan extension                               |
| `startupWarmups.mecab`                 | `true`  | Load the MeCab tokenizer                                                      |
| `startupWarmups.yomitanExtension`      | `true`  | Load the Yomitan extension                                                    |
| `startupWarmups.subtitleDictionaries`  | `true`  | Load the JLPT and frequency dictionaries                                      |
| `startupWarmups.jellyfinRemoteSession` | `false` | Connect the Jellyfin remote session (also needs Jellyfin remote auto-connect) |

### WebSocket server

Broadcasts plain subtitle text to external clients. See [WebSocket / Texthooker API](/websocket-texthooker-api) for payloads and client examples.

| Key                 | Default | What it does                                                                                                                   |
| ------------------- | ------- | ------------------------------------------------------------------------------------------------------------------------------ |
| `websocket.enabled` | `false` | `true`, `false`, or `"auto"` (start unless the [mpv_websocket](https://github.com/kuroahna/mpv_websocket) plugin is installed) |
| `websocket.port`    | `6677`  | Server port                                                                                                                    |

### Annotation WebSocket

A separate stream that adds token data (known word, N+1, frequency, JLPT, character names) to each subtitle. The bundled texthooker uses it.

| Key                           | Default | What it does                                            |
| ----------------------------- | ------- | ------------------------------------------------------- |
| `annotationWebsocket.enabled` | `false` | Start the annotated stream (independent of `websocket`) |
| `annotationWebsocket.port`    | `6678`  | Server port                                             |

### Texthooker

| Key                          | Default | What it does                                            |
| ---------------------------- | ------- | ------------------------------------------------------- |
| `texthooker.launchAtStartup` | `false` | Start the texthooker server when SubMiner starts        |
| `texthooker.openBrowser`     | `false` | Open the texthooker page in your browser when it starts |

## Subtitle display

### Subtitle style

Controls how primary and secondary subtitles look and which annotations they show. `css` and `secondary.css` take CSS declarations with normal property names. See [Subtitle annotations](/subtitle-annotations) for how known-word, N+1, frequency, JLPT, and character-name highlighting work.

```jsonc
{
  "subtitleStyle": {
    "css": { "font-size": "40px", "color": "#ffffff" },
    "secondary": { "css": { "font-size": "24px" } },
  },
}
```

| Key                                              | Default      | What it does                                                                                          |
| ------------------------------------------------ | ------------ | ----------------------------------------------------------------------------------------------------- |
| `subtitleStyle.primaryDefaultMode`               | `"visible"`  | Primary bar at startup: `hidden`, `visible`, or `hover`                                               |
| `subtitleStyle.css`                              | see example  | CSS for primary subtitles (font, size `35px`, color, shadow, and so on)                               |
| `subtitleStyle.secondary.css`                    | see example  | CSS for secondary subtitles (size `24px`)                                                             |
| `subtitleStyle.preserveLineBreaks`               | `false`      | Keep line breaks as mpv shows them instead of one line                                                |
| `subtitleStyle.autoPauseVideoOnHover`            | `true`       | Pause while the mouse is over subtitle text                                                           |
| `subtitleStyle.autoPauseVideoOnYomitanPopup`     | `true`       | Pause while a Yomitan popup is open                                                                   |
| `subtitleStyle.primaryVisibleOnYomitanPopup`     | `true`       | In hover mode, keep the primary bar visible while a popup is open                                     |
| `subtitleStyle.knownWordColor`                   | `#a6da95`    | Known-word highlight color                                                                            |
| `subtitleStyle.knownWordMaturityColors`          | see example  | `new`, `learning`, `young`, `mature` colors, used when `ankiConnect.knownWords.maturityEnabled` is on |
| `subtitleStyle.nPlusOneColor`                    | `#c6a0f6`    | N+1 target word color                                                                                 |
| `subtitleStyle.enableJlpt`                       | `false`      | Underline words by JLPT level                                                                         |
| `subtitleStyle.jlptColors`                       | see example  | Underline colors for `N1` to `N5`                                                                     |
| `subtitleStyle.nameMatchEnabled`                 | `false`      | Sync the character dictionary and color character names                                               |
| `subtitleStyle.nameMatchImagesEnabled`           | `false`      | Show small character portraits next to matched names                                                  |
| `subtitleStyle.nameMatchColor`                   | `#f5bde6`    | Character-name color                                                                                  |
| `subtitleStyle.frequencyDictionary.enabled`      | `false`      | Color words by frequency rank                                                                         |
| `subtitleStyle.frequencyDictionary.sourcePath`   | `""`         | Folder with `term_meta_bank_*.json` files. Empty searches the default locations                       |
| `subtitleStyle.frequencyDictionary.topX`         | `10000`      | Only color words ranked at or below this                                                              |
| `subtitleStyle.frequencyDictionary.mode`         | `"single"`   | `single` (one color) or `banded` (five colors, common to rare)                                        |
| `subtitleStyle.frequencyDictionary.matchMode`    | `"headword"` | Look up by `headword` (dictionary form) or `surface` (text as shown)                                  |
| `subtitleStyle.frequencyDictionary.singleColor`  | `#f5a97f`    | Color for `single` mode                                                                               |
| `subtitleStyle.frequencyDictionary.bandedColors` | see example  | Five colors for `banded` mode                                                                         |

Two CSS custom properties style the hovered word: `--subtitle-hover-token-color` (`#f4dbd6`) and `--subtitle-hover-token-background-color` (`transparent`). Set them inside `subtitleStyle.css`.

### Subtitle sidebar

A scrollable cue list for the current subtitle file. It only works when SubMiner could parse the active subtitle into cues. See [Subtitle sidebar](/subtitle-sidebar).

| Key                                 | Default       | What it does                                                                   |
| ----------------------------------- | ------------- | ------------------------------------------------------------------------------ |
| `subtitleSidebar.enabled`           | `true`        | Enable the sidebar                                                             |
| `subtitleSidebar.autoOpen`          | `false`       | Open it once when the overlay starts                                           |
| `subtitleSidebar.layout`            | `"overlay"`   | `overlay` floats over mpv. `embedded` reserves space on the right of the video |
| `subtitleSidebar.toggleKey`         | `"Backslash"` | `KeyboardEvent.code` that opens and closes it                                  |
| `subtitleSidebar.pauseVideoOnHover` | `true`        | Pause while hovering the cue list                                              |
| `subtitleSidebar.autoScroll`        | `true`        | Keep the active cue in view                                                    |
| `subtitleSidebar.css`               | see example   | CSS for the sidebar, plus the custom properties below                          |

Sidebar custom properties: `--subtitle-sidebar-max-width` (`420px`), `--subtitle-sidebar-timestamp-color`, `--subtitle-sidebar-active-line-color`, `--subtitle-sidebar-active-background-color`, `--subtitle-sidebar-hover-background-color`. Their defaults are in the example config.

If `embedded` layout places the video oddly on your system, switch back to `overlay`.

### Subtitle position

You can also drag subtitles with `Right-click + drag` while watching.

| Key                         | Default | What it does                                                     |
| --------------------------- | ------- | ---------------------------------------------------------------- |
| `subtitlePosition.yPercent` | `10`    | Starting distance from the bottom, as a percent of screen height |

### Secondary subtitles

Shows a second track, such as English, above the Japanese line.

Secondary subtitles do **not** auto-load by default (`autoLoadSecondarySub`, default: `false`). To load them for local and Jellyfin playback, turn it on and list the languages you want:

```json
{
  "secondarySub": {
    "secondarySubLanguages": ["eng", "en"],
    "autoLoadSecondarySub": true
  }
}
```

| Key                                  | Default   | What it does                                                                            |
| ------------------------------------ | --------- | --------------------------------------------------------------------------------------- |
| `secondarySub.secondarySubLanguages` | `[]`      | Language codes in priority order. Regular tracks win over Signs/Songs tracks            |
| `secondarySub.autoLoadSecondarySub`  | `false`   | Load a matching secondary track when the primary loads                                  |
| `secondarySub.defaultMode`           | `"hover"` | `hidden`, `visible` (always shown), or `hover` (shown when you hover the subtitle area) |

YouTube ignores the first two keys and always picks English. See [YouTube integration](/youtube-integration). `defaultMode` applies everywhere.

### Subtitle selection {#subtitle-selection}

Adds a modal for choosing mpv's primary and secondary subtitle tracks. Open it with `g` then `s` (`shortcuts.openSubtitleSelection`). While enabled, that shortcut replaces mpv's own binding for the same key. See [Keyboard shortcuts](/shortcuts) for sequence conflicts.

| Key                         | Default | What it does                        |
| --------------------------- | ------- | ----------------------------------- |
| `subtitleSelection.enabled` | `false` | Enable the subtitle selection modal |

## Keyboard and controls

### Keybindings

`keybindings` maps keys to mpv commands or SubMiner actions. Your entries merge with the defaults. The full default list is on [Keyboard shortcuts](/shortcuts).

```json
{
  "keybindings": [
    { "key": "Shift+ArrowRight", "command": ["seek", 30] },
    { "key": "MBTN_BACK", "command": ["sub-seek", -1] },
    { "key": "Space", "command": null }
  ]
}
```

- `key` uses `KeyboardEvent.code` names (`Space`, `KeyR`, `ArrowRight`) with optional `Ctrl+`, `Alt+`, `Shift+`, `Meta+`. Mouse buttons are `MBTN_LEFT`, `MBTN_MID`, `MBTN_RIGHT`, `MBTN_BACK`, `MBTN_FORWARD`.
- `command` is any mpv JSON IPC command array. Set it to `null` to disable a default.
- Commands starting with `__` run inside SubMiner: `__playlist-browser-open`, `__youtube-picker-open`, `__replay-subtitle`, `__play-next-subtitle`, `__runtime-options-open`, and `__runtime-option-cycle:<id>[:next|prev]`.
- Unused single-key bindings from your mpv config also work in the overlay. Your SubMiner bindings win on conflicts.

### Shortcuts configuration

`shortcuts` holds SubMiner's own actions (mining, copying, opening modals). Values are [Electron accelerator strings](https://www.electronjs.org/docs/latest/tutorial/keyboard-shortcuts) such as `"CommandOrControl+S"`. Set one to `null` to disable it. [Keyboard shortcuts](/shortcuts) lists every key, its default, and what it does. Anki shortcuts only run when `ankiConnect.enabled` is on.

| Key                            | Default | What it does                                             |
| ------------------------------ | ------- | -------------------------------------------------------- |
| `shortcuts.multiCopyTimeoutMs` | `3000`  | How long multi-copy and multi-mine wait for a digit (ms) |

### Controller support

Gamepad input for the overlay, through the browser Gamepad API. It only works while keyboard-only mode is on. Use the `Alt+C` modal to pick a controller and learn bindings, and `Alt+Shift+C` to see raw button and axis values. Default button actions are on [Keyboard shortcuts](/shortcuts).

| Key                                | Default  | What it does                                                                          |
| ---------------------------------- | -------- | ------------------------------------------------------------------------------------- |
| `controller.enabled`               | `false`  | Enable controller support. The `Alt+C` and `Alt+Shift+C` modals stay closed while off |
| `controller.smoothScroll`          | `true`   | Smooth popup scrolling                                                                |
| `controller.scrollPixelsPerSecond` | `900`    | Popup scroll speed                                                                    |
| `controller.horizontalJumpPixels`  | `160`    | Popup page-jump distance                                                              |
| `controller.stickDeadzone`         | `0.2`    | Stick deadzone                                                                        |
| `controller.triggerInputMode`      | `"auto"` | `auto`, `digital`, or `analog`. Use `analog` if your L2/R2 report analog values       |
| `controller.triggerDeadzone`       | `0.5`    | Trigger threshold for `auto` and `analog`                                             |
| `controller.repeatDelayMs`         | `320`    | Delay before a held button repeats                                                    |
| `controller.repeatIntervalMs`      | `120`    | Repeat interval for held buttons                                                      |

Bindings are set with `Alt+C` learn mode, which saves them per controller.

## Anki integration

### AnkiConnect

Creates and updates Anki cards with sentence, audio, and screenshot. Needs the [AnkiConnect](https://github.com/FooSoft/anki-connect) add-on and ffmpeg. See [Anki integration](/anki-integration) for setup, the proxy, and media options in detail.

```json
{
  "ankiConnect": {
    "deck": "Mining",
    "fields": { "audio": "SentenceAudio", "image": "Picture" },
    "knownWords": { "highlightEnabled": true, "decks": { "Mining": ["Expression"] } }
  }
}
```

**Connection**

| Key                             | Default                   | What it does                                                                   |
| ------------------------------- | ------------------------- | ------------------------------------------------------------------------------ |
| `ankiConnect.enabled`           | `true`                    | Enable Anki integration                                                        |
| `ankiConnect.url`               | `"http://127.0.0.1:8765"` | AnkiConnect URL                                                                |
| `ankiConnect.pollingRate`       | `3000`                    | Milliseconds between checks for new cards (polling mode)                       |
| `ankiConnect.proxy.enabled`     | `true`                    | Run a local AnkiConnect proxy so cards added through it are updated right away |
| `ankiConnect.proxy.host`        | `"127.0.0.1"`             | Proxy bind host                                                                |
| `ankiConnect.proxy.port`        | `8766`                    | Proxy bind port                                                                |
| `ankiConnect.proxy.upstreamUrl` | `"http://127.0.0.1:8765"` | Where the proxy forwards requests                                              |
| `ankiConnect.tags`              | `["SubMiner"]`            | Tags added to mined and updated cards. `[]` disables                           |
| `ankiConnect.deck`              | `""`                      | Deck for duplicate checks and enrichment. Empty uses Yomitan's mining deck     |

**Fields**

| Key                            | Default                | What it does                                                                                                                   |
| ------------------------------ | ---------------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| `ankiConnect.fields.word`      | `"Expression"`         | Word field                                                                                                                     |
| `ankiConnect.fields.audio`     | `"ExpressionAudio"`    | Field that receives sentence audio. Set a separate field such as `SentenceAudio` so it does not overwrite Yomitan's word audio |
| `ankiConnect.fields.wordAudio` | `"ExpressionAudio"`    | Existing word-audio field, read only to time animated images                                                                   |
| `ankiConnect.fields.image`     | `"Picture"`            | Screenshot field                                                                                                               |
| `ankiConnect.fields.sentence`  | `"Sentence"`           | Sentence field                                                                                                                 |
| `ankiConnect.fields.miscInfo`  | `"MiscInfo"`           | Metadata field. `null` disables                                                                                                |
| `ankiConnect.metadata.pattern` | `"[SubMiner] %f (%t)"` | MiscInfo template: `%f` filename, `%F` filename with extension, `%t` time, `%T` time with ms, `<br>` newline                   |

**Media**

| Key                                              | Default    | What it does                                                 |
| ------------------------------------------------ | ---------- | ------------------------------------------------------------ |
| `ankiConnect.media.generateAudio`                | `true`     | Cut a sentence audio clip                                    |
| `ankiConnect.media.generateImage`                | `true`     | Capture a screenshot or animation                            |
| `ankiConnect.media.imageType`                    | `"static"` | `static` or `avif` (animated)                                |
| `ankiConnect.media.imageFormat`                  | `"jpg"`    | Static format: `jpg`, `png`, `webp`                          |
| `ankiConnect.media.imageQuality`                 | `92`       | JPG/WebP quality. PNG ignores it                             |
| `ankiConnect.media.imageMaxWidth`                | `0`        | Max static width in px. `0` keeps the source size            |
| `ankiConnect.media.imageMaxHeight`               | `0`        | Max static height in px. `0` keeps the source size           |
| `ankiConnect.media.animatedFps`                  | `10`       | AVIF frame rate                                              |
| `ankiConnect.media.animatedMaxWidth`             | `640`      | AVIF max width                                               |
| `ankiConnect.media.animatedMaxHeight`            | `0`        | AVIF max height. `0` keeps the aspect ratio                  |
| `ankiConnect.media.animatedCrf`                  | `35`       | AVIF quality. Lower is better and larger                     |
| `ankiConnect.media.syncAnimatedImageToWordAudio` | `true`     | Hold the first AVIF frame for the length of the word audio   |
| `ankiConnect.media.normalizeAudio`               | `true`     | Normalize clip loudness                                      |
| `ankiConnect.media.mirrorMpvVolume`              | `true`     | Apply mpv's current volume to the clip                       |
| `ankiConnect.media.reviewTiming`                 | `false`    | Pause and let you adjust clip timing before media is created |
| `ankiConnect.media.audioPadding`                 | `0`        | Seconds added to both ends of audio and AVIF clips           |
| `ankiConnect.media.fallbackDuration`             | `3`        | Clip length in seconds when subtitle timing is missing       |
| `ankiConnect.media.maxMediaDuration`             | `30`       | Longest allowed clip in seconds. `0` removes the cap         |

**Behavior**

| Key                                       | Default     | What it does                                                             |
| ----------------------------------------- | ----------- | ------------------------------------------------------------------------ |
| `ankiConnect.behavior.autoUpdateNewCards` | `true`      | Fill new cards automatically. When off, use the manual shortcuts         |
| `ankiConnect.behavior.overwriteAudio`     | `true`      | Replace existing audio. When off, add alongside it                       |
| `ankiConnect.behavior.overwriteImage`     | `true`      | Replace existing images. When off, add alongside them                    |
| `ankiConnect.behavior.mediaInsertMode`    | `"append"`  | `append` or `prepend` when not overwriting                               |
| `ankiConnect.behavior.highlightWord`      | `true`      | Bold the mined word in the sentence field                                |
| `ankiConnect.behavior.notificationType`   | `"overlay"` | Where mining and status messages go: `overlay`, `system`, `both`, `none` |

**Known words and N+1**

| Key                                               | Default      | What it does                                                                     |
| ------------------------------------------------- | ------------ | -------------------------------------------------------------------------------- |
| `ankiConnect.knownWords.highlightEnabled`         | `false`      | Highlight words that already exist in your Anki decks                            |
| `ankiConnect.knownWords.decks`                    | `{}`         | Decks and word fields to read, for example `{ "Kaishi 1.5k": ["Word"] }`         |
| `ankiConnect.knownWords.matchMode`                | `"headword"` | Match by `headword` or `surface` text                                            |
| `ankiConnect.knownWords.refreshMinutes`           | `1440`       | Minutes between cache refreshes                                                  |
| `ankiConnect.knownWords.addMinedWordsImmediately` | `true`       | Add newly mined words to the cache right away                                    |
| `ankiConnect.knownWords.maturityEnabled`          | `false`      | Color known words by card maturity using `subtitleStyle.knownWordMaturityColors` |
| `ankiConnect.knownWords.matureThresholdDays`      | `21`         | Interval in days at which a card counts as mature                                |
| `ankiConnect.nPlusOne.enabled`                    | `false`      | Highlight the only unknown word in a sentence. Needs known-word data             |
| `ankiConnect.nPlusOne.minSentenceWords`           | `3`          | Minimum words in a sentence before N+1 applies                                   |

Use word fields such as `Expression` or `Word` in `knownWords.decks`, not reading fields. See [Subtitle annotations](/subtitle-annotations) for how matching and maturity tiers work.

### Kiku/Lapis integration {#kiku-lapis-integration}

Note-type behavior for [Lapis](https://github.com/donkuri/lapis), [Kiku](https://kiku.youyoumu.my.id/), and [Senren](https://github.com/BrenoAqua/Senren). With both Lapis and Kiku on, Kiku handles duplicates and the sentence-card model comes from `isLapis`. Kiku and Senren are mutually exclusive. If both are on, Kiku wins and SubMiner logs a warning. See [Anki integration](/anki-integration) for details.

| Key                                          | Default               | What it does                                           |
| -------------------------------------------- | --------------------- | ------------------------------------------------------ |
| `ankiConnect.isLapis.enabled`                | `false`               | Mine dedicated sentence cards (`IsSentenceCard`)       |
| `ankiConnect.isLapis.sentenceCardModel`      | `"Lapis"`             | Note type used for sentence cards                      |
| `ankiConnect.isKiku.enabled`                 | `false`               | Merge duplicate word cards                             |
| `ankiConnect.isKiku.fieldGrouping`           | `"disabled"`          | `auto`, `manual`, or `disabled`. See below             |
| `ankiConnect.isKiku.deleteDuplicateInAuto`   | `true`                | Delete the duplicate after an `auto` merge             |
| `ankiConnect.isSenren.enabled`               | `false`               | Merge duplicates using Senren's scene-switching format |
| `ankiConnect.isSenren.fieldGrouping`         | `"auto"`              | `auto`, `manual`, or `disabled`                        |
| `ankiConnect.isSenren.deleteDuplicateInAuto` | `true`                | Delete the duplicate after an `auto` merge             |
| `ankiConnect.lapisKiku.wordCardKind`         | `"word-and-sentence"` | Card-type flag set on word cards. See below            |

### Word card type

When SubMiner fills the sentence on a word card, it sets one card-type flag and clears the others. Only applies while `isLapis` or `isKiku` is on. Cards from Mine Sentence and Mine Audio keep their own flag.

| `wordCardKind`                | Flag set                |
| ----------------------------- | ----------------------- |
| `word-and-sentence` (default) | `IsWordAndSentenceCard` |
| `click`                       | `IsClickCard`           |
| `sentence`                    | `IsSentenceCard`        |
| `audio`                       | `IsAudioCard`           |
| `none`                        | none, flags left as-is  |

### Field grouping modes

| Mode       | What happens when you mine a duplicate                                                                     |
| ---------- | ---------------------------------------------------------------------------------------------------------- |
| `auto`     | Merges the new card into the existing one. `deleteDuplicateInAuto` decides whether the new card is deleted |
| `manual`   | Pauses playback and opens a dialog to choose which card to keep and whether to delete the other            |
| `disabled` | Leaves both cards as they are                                                                              |

<video controls playsinline preload="metadata" :poster="withBase('/assets/kiku-integration-poster.jpg')" style="width: 100%; max-width: 960px;">
  <source :src="withBase('/assets/kiku-integration.webm')" type="video/webm" />
  <source :src="withBase('/assets/kiku-integration.mp4')" type="video/mp4" />
  Your browser does not support the video tag.
</video>

## External integrations

### Jimaku

Search and download Japanese subtitles from [Jimaku](https://jimaku.cc). See [Jimaku integration](/jimaku-integration).

| Key                         | Default               | What it does                                               |
| --------------------------- | --------------------- | ---------------------------------------------------------- |
| `jimaku.apiKey`             | `""`                  | API key. Optional, but raises your rate limit              |
| `jimaku.apiKeyCommand`      | `""`                  | Shell command that prints the key. Use instead of `apiKey` |
| `jimaku.apiBaseUrl`         | `"https://jimaku.cc"` | API base URL                                               |
| `jimaku.languagePreference` | `"ja"`                | Preferred language: `ja`, `en`, or `none`                  |
| `jimaku.maxEntryResults`    | `10`                  | Maximum search results                                     |

### TsukiHime

Subtitle search that needs no account or key. It does need `xz` on your `PATH`. The shortcut is `shortcuts.openTsukihime`. See [TsukiHime integration](/tsukihime-integration).

| Key                          | Default                          | What it does                                         |
| ---------------------------- | -------------------------------- | ---------------------------------------------------- |
| `tsukihime.apiBaseUrl`       | `"https://api.tsukihime.org/v1"` | API base URL. Only change it for a mirror            |
| `tsukihime.maxSearchResults` | `10`                             | Maximum releases per search (the API caps it at 100) |

### TMDB

Posters, synopses, and show grouping for live-action titles in the stats [Library](/immersion-tracking). Release builds include a TMDB key, so you only need your own to use your own quota or when running from source. Get one free under **Settings > API** on [themoviedb.org](https://www.themoviedb.org/settings/api). Either the API key or the read access token works.

| Key                  | Default | What it does                                               |
| -------------------- | ------- | ---------------------------------------------------------- |
| `tmdb.apiKey`        | `""`    | Your TMDB key or token. Overrides the bundled key          |
| `tmdb.apiKeyCommand` | `""`    | Shell command that prints the key. Use instead of `apiKey` |

This product uses the TMDB API but is not endorsed or certified by TMDB.

### Japanese subtitle generation

Transcribes Japanese subtitles locally with whisper.cpp. Open it with `Ctrl+Shift+G` (`shortcuts.openSubtitleGeneration`) or from the subtitle sidebar. See [Subtitle generation](/subtitle-generation).

| Key                               | Default   | What it does                                                            |
| --------------------------------- | --------- | ----------------------------------------------------------------------- |
| `subtitleGeneration.modelPath`    | `""`      | Path to a multilingual whisper.cpp GGML model. Overrides `managedModel` |
| `subtitleGeneration.managedModel` | `"small"` | Model SubMiner downloads and uses when `modelPath` is empty             |
| `subtitleGeneration.threads`      | `4`       | CPU threads                                                             |
| `subtitleGeneration.vadModelPath` | `""`      | Silero VAD model. Set it to focus on spoken dialogue by default         |
| `subtitleGeneration.whisperPath`  | `""`      | `whisper-cli` path. Empty searches `PATH`                               |
| `subtitleGeneration.vadPath`      | `""`      | Speech detector path. Empty searches `PATH`                             |
| `subtitleGeneration.ffmpegPath`   | `""`      | `ffmpeg` path. Empty searches `PATH`                                    |
| `subtitleGeneration.ffprobePath`  | `""`      | `ffprobe` path. Empty searches `PATH`                                   |

### Subtitle sync

Retimes a subtitle track with [`alass`](https://github.com/kaegi/alass) (against another subtitle or the video) or [`ffsubsync`](https://github.com/smacke/ffsubsync) (against the video's audio). Install them yourself. Open the picker with `Ctrl+Alt+S` (`shortcuts.triggerSubsync`).

| Key                      | Default | What it does                                                        |
| ------------------------ | ------- | ------------------------------------------------------------------- |
| `subsync.alass_path`     | `""`    | `alass` path. Empty uses `/usr/bin/alass`                           |
| `subsync.ffsubsync_path` | `""`    | `ffsubsync` path. Empty uses `/usr/bin/ffsubsync`                   |
| `subsync.ffmpeg_path`    | `""`    | `ffmpeg` path. Empty uses `/usr/bin/ffmpeg`                         |
| `subsync.replace`        | `true`  | Overwrite the subtitle file. When off, write `<name>_retimed.<ext>` |

If a tool lives somewhere else, such as on macOS or Windows, set its path.

### AniList

Updates your AniList watch progress after an episode, and controls the character dictionary. With `enabled` on and no token, SubMiner opens a login window. See [AniList integration](/anilist-integration) and [Character dictionary](/character-dictionary).

| Key                                                                    | Default | What it does                                                  |
| ---------------------------------------------------------------------- | ------- | ------------------------------------------------------------- |
| `anilist.enabled`                                                      | `false` | Enable progress updates                                       |
| `anilist.accessToken`                                                  | `""`    | Token override. Empty uses the token saved during login       |
| `anilist.characterDictionary.maxLoaded`                                | `3`     | How many recent shows stay in the merged character dictionary |
| `anilist.characterDictionary.collapsibleSections.description`          | `false` | Open the Description section by default                       |
| `anilist.characterDictionary.collapsibleSections.characterInformation` | `false` | Open the Character Information section by default             |
| `anilist.characterDictionary.collapsibleSections.voicedBy`             | `false` | Open the Voiced by section by default                         |

### Yomitan

Point SubMiner at another app's Yomitan Electron profile to reuse its dictionaries and settings. For GameSentenceMiner on Linux this is usually `~/.config/gsm_overlay`.

| Key                           | Default | What it does                                                            |
| ----------------------------- | ------- | ----------------------------------------------------------------------- |
| `yomitan.externalProfilePath` | `""`    | Absolute or `~` path to the external profile. Empty uses SubMiner's own |

In external-profile mode, SubMiner only reads the profile. It does not open its own Yomitan settings, does not change dictionaries, and turns off all character-dictionary features.

### Jellyfin

Log in to a Jellyfin server, browse libraries, and play or cast to SubMiner. Login tokens are stored encrypted, not in this file. See [Jellyfin integration](/jellyfin-integration).

| Key                                 | Default                          | What it does                                    |
| ----------------------------------- | -------------------------------- | ----------------------------------------------- |
| `jellyfin.enabled`                  | `false`                          | Enable Jellyfin                                 |
| `jellyfin.serverUrl`                | `""`                             | Server URL, for example `http://localhost:8096` |
| `jellyfin.username`                 | `""`                             | Default username for `subminer jellyfin -l`     |
| `jellyfin.remoteControlEnabled`     | `true`                           | Let Jellyfin apps cast to SubMiner              |
| `jellyfin.remoteControlAutoConnect` | `true`                           | Connect the cast session on startup             |
| `jellyfin.autoAnnounce`             | `false`                          | Announce SubMiner as a cast target on connect   |
| `jellyfin.pullPictures`             | `false`                          | Fetch posters for launcher pickers              |
| `jellyfin.iconCacheDir`             | `"/tmp/subminer-jellyfin-icons"` | Poster cache folder                             |
| `jellyfin.directPlayPreferred`      | `true`                           | Try direct play before transcoding              |
| `jellyfin.transcodeVideoCodec`      | `"h264"`                         | Codec requested when transcoding                |

### Discord rich presence

Shows what you are watching on your Discord profile. Needs the Discord desktop app running. If Discord is closed, SubMiner skips updates.

| Key                                | Default     | What it does                                                       |
| ---------------------------------- | ----------- | ------------------------------------------------------------------ |
| `discordPresence.enabled`          | `true`      | Enable rich presence                                               |
| `discordPresence.presenceStyle`    | `"default"` | Card text: `default`, `meme`, `japanese` (all Japanese), `minimal` |
| `discordPresence.updateIntervalMs` | `3000`      | Minimum ms between updates                                         |
| `discordPresence.debounceMs`       | `750`       | Debounce for bursts of playback events                             |

### Immersion tracking

Records watch sessions, subtitle lines, and mining in a local SQLite database that feeds the stats dashboard. See [Immersion tracking](/immersion-tracking) for retention and storage details. To turn it off for one run, start with `SUBMINER_DISABLE_IMMERSION_TRACKING=1 subminer`.

| Key                                              | Default      | What it does                                                      |
| ------------------------------------------------ | ------------ | ----------------------------------------------------------------- |
| `immersionTracking.enabled`                      | `true`       | Enable tracking                                                   |
| `immersionTracking.dbPath`                       | `""`         | Database path. Empty uses `immersion.sqlite` in the config folder |
| `immersionTracking.batchSize`                    | `25`         | Writes per transaction                                            |
| `immersionTracking.flushIntervalMs`              | `500`        | Maximum ms before queued writes are saved                         |
| `immersionTracking.queueCap`                     | `1000`       | Queue size. The oldest writes drop when full                      |
| `immersionTracking.payloadCapBytes`              | `256`        | Maximum event payload size before truncation                      |
| `immersionTracking.maintenanceIntervalMs`        | `86400000`   | How often pruning and rollups run (24 h)                          |
| `immersionTracking.retentionMode`                | `"preset"`   | `preset` uses `retentionPreset`. `advanced` uses `retention.*`    |
| `immersionTracking.retentionPreset`              | `"balanced"` | `minimal`, `balanced`, or `deep-history`                          |
| `immersionTracking.retention.eventsDays`         | `0`          | Days to keep raw events. `0` keeps everything                     |
| `immersionTracking.retention.telemetryDays`      | `0`          | Days to keep telemetry                                            |
| `immersionTracking.retention.sessionsDays`       | `0`          | Days to keep sessions                                             |
| `immersionTracking.retention.dailyRollupsDays`   | `0`          | Days to keep daily rollups                                        |
| `immersionTracking.retention.monthlyRollupsDays` | `0`          | Days to keep monthly rollups                                      |
| `immersionTracking.retention.vacuumIntervalDays` | `0`          | Days between `VACUUM` runs. `0` disables                          |
| `immersionTracking.lifetimeSummaries.global`     | `true`       | Keep all-time totals                                              |
| `immersionTracking.lifetimeSummaries.anime`      | `true`       | Keep per-show totals                                              |
| `immersionTracking.lifetimeSummaries.media`      | `true`       | Keep per-file totals                                              |

### Stats dashboard

A local web dashboard at `http://127.0.0.1:<serverPort>`, also available as an overlay inside SubMiner. It reads the immersion tracking database, so tracking must be on. See [Immersion tracking](/immersion-tracking).

| Key                     | Default       | What it does                                                       |
| ----------------------- | ------------- | ------------------------------------------------------------------ |
| `stats.toggleKey`       | `"Backquote"` | Key that toggles the stats overlay (overlay focus only)            |
| `stats.markWatchedKey`  | `"KeyW"`      | Key that marks the video watched and plays the next playlist entry |
| `stats.serverPort`      | `6969`        | Dashboard port                                                     |
| `stats.autoStartServer` | `true`        | Start the dashboard server once tracking is active                 |
| `stats.autoOpenBrowser` | `false`       | Open the browser when `subminer stats` starts the server           |

### MPV launcher

Settings for mpv instances that SubMiner starts, and for the bundled mpv plugin. See [mpv plugin](/mpv-plugin).

| Key                          | Default           | What it does                                                                |
| ---------------------------- | ----------------- | --------------------------------------------------------------------------- |
| `mpv.executablePath`         | `""`              | Path to `mpv.exe` on Windows. Empty checks `SUBMINER_MPV_PATH`, then `PATH` |
| `mpv.launchMode`             | `"normal"`        | Window state: `normal`, `maximized`, or `fullscreen`                        |
| `mpv.profile`                | `""`              | mpv profile to pass. Combined with a launcher `--profile` if both are set   |
| `mpv.socketPath`             | platform-specific | mpv IPC socket. See the warning under [Config file](#configuration-file)    |
| `mpv.backend`                | `"auto"`          | Window tracking: `auto`, `hyprland`, `sway`, `x11`, `macos`, `windows`      |
| `mpv.autoStartSubMiner`      | `true`            | Start SubMiner in the background when mpv loads a file                      |
| `mpv.pauseUntilOverlayReady` | `true`            | Keep mpv paused until subtitles are ready, up to 30 seconds                 |
| `mpv.subminerBinaryPath`     | `""`              | SubMiner app path for the plugin. Empty uses the detected path              |
| `mpv.aniskipEnabled`         | `true`            | Detect intros with AniSkip and show a skip prompt                           |
| `mpv.aniskipButtonKey`       | `"TAB"`           | mpv key that skips the intro while the prompt is shown                      |

### YouTube playback settings

Language and card-media settings for YouTube playback. YouTube always loads a Japanese primary and English secondary track, preferring manual uploads over auto captions. See [YouTube integration](/youtube-integration).

| Key                            | Default         | What it does                                                                                 |
| ------------------------------ | --------------- | -------------------------------------------------------------------------------------------- |
| `youtube.primarySubLanguages`  | `["ja", "jpn"]` | Languages that count as a valid primary track, also used for local playback                  |
| `youtube.mediaCache.mode`      | `"direct"`      | `direct` cuts card media from the stream. `background` downloads the video with yt-dlp first |
| `youtube.mediaCache.maxHeight` | `720`           | Maximum download height in `background` mode. `0` is unlimited                               |

Use `background` if card media fails with YouTube `403` errors. Cards mined before the download finishes get their text right away and their audio and image once the file is ready.
