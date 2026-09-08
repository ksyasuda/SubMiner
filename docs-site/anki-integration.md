# Anki integration

SubMiner uses the [AnkiConnect](https://ankiweb.net/shared/info/2055492159) add-on to create and update Anki cards with sentence context, audio, and screenshots.
This project is built primarily for [Kiku](https://kiku.youyoumu.my.id/) and [Lapis](https://github.com/donkuri/lapis) note types, including sentence-card and field-grouping behavior.

::: tip New to these terms?

- **Anki** is the flashcard app where your study cards live.
- **AnkiConnect** is a free add-on that lets other programs (like SubMiner) talk to Anki over a local connection. SubMiner needs it installed to add or edit cards.
- A **note type** (also called a "model") is the template that defines what a card looks like - for example the Kiku or Lapis templates many Japanese learners use.
- A **field** is one labeled slot in that template, such as `Sentence`, `Expression`, or `Picture`. SubMiner fills these fields when it mines a card.
  :::

## Prerequisites

1. Install [Anki](https://apps.ankiweb.net/).
2. Install the [AnkiConnect](https://ankiweb.net/shared/info/2055492159) add-on (code: `2055492159`).
3. Keep Anki running while using SubMiner.

AnkiConnect listens on `http://127.0.0.1:8765` by default. If you changed the port in AnkiConnect's settings, update `ankiConnect.url` in your SubMiner config.

## Auto-enrichment transport

When you add a word via Yomitan, SubMiner detects the new card and fills in the sentence, audio, and image fields automatically. Two detection methods are available:

**Proxy mode** (default) - SubMiner runs a small local server between Yomitan and Anki. Yomitan sends the new card to SubMiner, SubMiner fills in the media fields, and the finished card goes on to Anki. There is no polling delay.

**Polling mode** (fallback, when the proxy is disabled) - SubMiner asks AnkiConnect every few seconds whether new cards showed up, then fills them in. Less to configure, at the cost of roughly a 3 second delay.

Use proxy mode unless your Yomitan runs in a browser rather than the bundled instance, in which case polling is the simpler path.

In both modes, the enrichment workflow is the same:

1. Checks if a duplicate expression already exists (for field grouping).
2. Updates the sentence field with the current subtitle.
3. Generates and uploads audio and image media.
4. Writes metadata to the miscInfo field.

Polling mode uses the query `"deck:<ankiConnect.deck>" added:1` to find recently added cards. If no deck is configured, it searches all decks (`added:1`). In Settings, the AnkiConnect deck dropdown auto-fills and persists Yomitan's current mining deck when available, then falls back to the decks reported by AnkiConnect; stats-dashboard mining also falls back to Yomitan's mining deck when `ankiConnect.deck` is empty.
Known-word sync scope is controlled by `ankiConnect.knownWords.decks`.

### Proxy mode setup (Yomitan / texthooker)

```jsonc
"ankiConnect": {
  "url": "http://127.0.0.1:8765", // real AnkiConnect
  "proxy": {
    "enabled": true,
    "host": "127.0.0.1",
    "port": 8766,
    "upstreamUrl": "http://127.0.0.1:8765"
  }
}
```

Then point Yomitan/clients to `http://127.0.0.1:8766` instead of `8765`.

When SubMiner loads the bundled Yomitan extension, it also attempts to update the **currently active Yomitan profile**'s Anki server to the active SubMiner endpoint (falling back to `profiles[0]` if the active-profile index is invalid):

- proxy URL when `ankiConnect.proxy.enabled` is `true`
- direct `ankiConnect.url` when proxy mode is disabled

To avoid clobbering custom setups, this auto-update only changes the profile when its current server is blank or the stock Yomitan default (`http://127.0.0.1:8765`).

For browser-based Yomitan or other external clients (for example Texthooker in a normal browser profile), set their Anki server to the same proxy URL separately: `http://127.0.0.1:8766` (or your configured `proxy.host` + `proxy.port`).

### Browser/Yomitan external setup (separate profile)

If you want SubMiner to use proxy mode without touching your main/default Yomitan profile, create or select a separate Yomitan profile just for SubMiner and set its Anki server to the proxy URL.

That profile isolation gives you both benefits:

- SubMiner can auto-enrich immediately via proxy.
- Your default Yomitan profile keeps its existing Anki server setting.

In Yomitan, go to Settings → Profile and:

1. Create a profile for SubMiner (or choose one dedicated profile).
2. Open Anki settings for that profile.
3. Set server to `http://127.0.0.1:8766` (or your configured proxy URL).
4. Save and make that profile active when using SubMiner.

This is only for non-bundled, external/browser Yomitan or other clients. The bundled profile auto-update logic only targets the active profile when its server is blank or still default.

### Proxy troubleshooting (quick checks)

If auto-enrichment appears to do nothing:

1. Confirm proxy listener is running while SubMiner is active:

```bash
ss -ltnp | rg 8766
```

2. Confirm requests can pass through the proxy:

```bash
curl -sS http://127.0.0.1:8766 \
  -H 'content-type: application/json' \
  -d '{"action":"version","version":2}'
```

3. Check the log sinks in `~/.config/SubMiner/logs/`:

- App runtime log: `app-YYYY-MM-DD.log`
- Launcher log: `launcher-YYYY-MM-DD.log`
- mpv log: `mpv-YYYY-MM-DD.log`

4. Check that the config JSONC parses and the logging shape is right:

```jsonc
"logging": {
  "level": "debug"
}
```

`"logging": "debug"` is invalid for current schema and can break reload/start behavior.

## Field mapping

SubMiner maps its data to your Anki note fields. Configure these under `ankiConnect.fields`:

```jsonc
"ankiConnect": {
  "fields": {
    "word": "Expression",        // mined word / expression text
    "audio": "SentenceAudio",    // sentence audio clip cut from the video
    "image": "Picture",          // screenshot or animated clip
    "sentence": "Sentence",      // subtitle text
    "miscInfo": "MiscInfo"       // metadata (filename, timestamp)
  }
}
```

`fields.audio` receives the **sentence** audio SubMiner cuts from the video, not word audio. Yomitan writes its own dictionary audio when you mine, so point this at a separate field such as `SentenceAudio` to keep the two apart. The built-in default is still `ExpressionAudio`, which collides with Yomitan on note types that use that field for word audio.

Field names are matched against your Anki note type case-insensitively (an exact match wins, then a lowercase comparison). If a configured field does not exist on the note type, SubMiner skips it without error.

These mappings always control normal word-card enrichment, including Yomitan proxy/polling updates and manual clipboard updates. Enabling Lapis or Kiku does not replace the configured word-card sentence and audio fields with `Sentence` and `SentenceAudio`. The dedicated sentence-card and audio-card shortcuts still use those Lapis/Kiku field names.

Two related options live alongside `fields`: `ankiConnect.deck` (target deck; empty falls back as described above) and `ankiConnect.tags` (tags added to mined cards, default `["SubMiner"]`; set `[]` to disable tagging). The `miscInfo` content is controlled by `ankiConnect.metadata.pattern` (default `[SubMiner] %f (%t)`; tokens: `%f` filename, `%F` filename with extension, `%t` timestamp, `%T` timestamp with milliseconds, `<br>` newline).

### Minimal config

If you only want sentence and audio on your cards:

```jsonc
"ankiConnect": {
  "enabled": true,
  "fields": {
    "sentence": "Sentence",
    "audio": "SentenceAudio"
  }
}
```

## Media generation

SubMiner shells out to FFmpeg for audio clips and screenshots, so FFmpeg has to be installed and on `PATH`.

For remote streams such as Jellyfin playback, SubMiner downloads the clip's time window once into a temporary Matroska file (a stream copy, no re-encoding) and reads the timing review waveform, audio preview, audio, and image from that file instead of fetching the stream again for each step. The window covers the clip plus padding, plus the visible timeline in timing review, and grows when you reveal more of the timeline. It is deleted when a different window replaces it, after ten minutes without use, or when SubMiner exits. If the download fails, media generation reads the remote stream directly as before.

### Audio

Audio is extracted from the video file using the subtitle's start and end timestamps. Padding is opt-in; keep it at `0` when you want sentence audio to start exactly at the mined sentence.

```jsonc
"ankiConnect": {
  "media": {
    "generateAudio": true,
    "normalizeAudio": true,      // normalize generated clip loudness
    "mirrorMpvVolume": true,     // apply the current mpv volume level
    "reviewTiming": false,       // review and adjust timing before media generation
    "audioPadding": 0,           // optional seconds before and after subtitle timing
    "maxMediaDuration": 30       // cap total duration in seconds
  }
}
```

Output format: MP3 at 44100 Hz. If the video has multiple audio streams, SubMiner uses the active stream. Generated sentence audio is loudness-normalized to -23 LUFS by default during extraction; set `normalizeAudio` to `false` to keep raw source loudness. When subtitle timing is missing, clips fall back to `media.fallbackDuration` seconds (default `3`). Changing these settings applies to the next extraction without restarting SubMiner.

`mirrorMpvVolume` is also enabled by default. Immediately before extracting each playback-overlay card's audio, SubMiner reads mpv's numeric `volume` and applies mpv's cubic software-volume curve after loudness normalization. For example, mpv volume `50` produces `0.5³ = 0.125` gain. Amplified output above mpv volume `100` is limited to a `-1 dBFS` ceiling before MP3 encoding to prevent clipping. It ignores mpv's separate `mute` state. If the volume property is missing, invalid, or unavailable, extraction continues with unity scaling; disabling this option skips the query and volume filter. Changing this setting applies to the next extraction without restarting SubMiner. YouTube cards queued for a background media-cache download retain the volume captured when the card was mined. Stats-dashboard mining does not currently have access to the active mpv property client, so it does not apply mpv volume scaling.

The audio is uploaded to Anki's media folder and inserted as `[sound:audio_<timestamp>.mp3]`.

Set `media.reviewTiming` to `true` to pause playback and check the clip before its media is generated. It applies to word, sentence, and audio cards.

The review opens on the subtitle range plus your configured audio padding. Subtitles usually hang around after the dialogue has stopped, so once the waveform loads, an untouched clip end pulls back to just after the last speech in the line. The Line end rail still marks the original subtitle timing, Reset puts it back, and a line whose speech runs right through its end is left alone.

**Adjusting the clip.** Drag either edge to trim, drag the middle to slide the whole clip without changing its length, or click anywhere on the waveform to snap the nearer edge there. A focused edge also moves with the arrow keys: 100 ms per press, or 500 ms with Shift. The 100 ms buttons do the same thing. Earlier and Later each reveal two more seconds of timeline without moving the selection.

**Keys.** Space previews the selection with a playhead sweeping the clip. The preview ends when the hidden player has actually played the last sample, so Bluetooth output latency does not clip the tail. Enter confirms and Escape cancels.

**The waveform.** SubMiner reads a center channel when one carries dialogue and falls back to a mono mix otherwise, keeps only the 250 to 3500 Hz speech band, and draws each slice's loudness against the clip's own noise floor. Steady background music flattens out and dialogue stands up, which makes it much easier to tell adjacent lines apart. The mined subtitle appears as a tinted band with labeled line-start and line-end rails. If waveform analysis fails, the timing controls still work.

The range you confirm is used exactly as-is; SubMiner does not add audio padding a second time. Static screenshots take its midpoint, and animated AVIF clips cover the whole range.

**Pulling in adjacent lines.** Press `P` or `N`, or use the Prev and Next steppers above the sentence preview, to add the previous or next subtitle line. Repeat for as many lines as exist. Shift+`P` and Shift+`N` remove them again. The sentence preview lists every included line with the mined one highlighted, so you always see the sentence field before confirming. The clip bounds and the waveform rails follow the outermost added line, keeping the review's audio padding.

Confirming writes the combined lines to the sentence field. Reset drops the added lines along with any timing changes. Adjacent lines come from the parsed subtitle track when one is loaded; otherwise you only get lines that already played. A clip capped by `media.maxMediaDuration` still keeps the full combined sentence even when the audio cannot stretch to cover every added line.

**Canceling.** You can go back to editing, finish with the original timing, create the card without audio or an image, or discard it. Discard deletes an existing Yomitan or audio card, and skips creation entirely for a direct sentence card. A failed audio preview does not block confirmation or card creation.

Clipboard updates and stats-dashboard mining never open timing review. The option is off by default and hot-reloads. **Review Media Timing** in the runtime options palette (`Ctrl/Cmd+Shift+O`) toggles it for the current session.

### Screenshots (static)

A single frame is captured at the current playback position.

```jsonc
"ankiConnect": {
  "media": {
    "generateImage": true,
    "imageType": "static",
    "imageFormat": "jpg",        // "jpg", "png", or "webp"
    "imageQuality": 92,          // 1–100
    "imageMaxWidth": 0,          // 0 = preserve source resolution
    "imageMaxHeight": 0
  }
}
```

### Animated clips (AVIF)

SubMiner can produce an animated AVIF spanning the subtitle duration instead of a still frame.

```jsonc
"ankiConnect": {
  "media": {
    "generateImage": true,
    "imageType": "avif",
    "animatedFps": 10,
    "animatedMaxWidth": 640,
    "animatedMaxHeight": 0,      // 0 = preserve aspect ratio
    "animatedCrf": 35            // 0–63, lower = better quality
  }
}
```

Animated AVIF requires an AV1 encoder (`libaom-av1`, `libsvtav1`, or `librav1e`) in your FFmpeg build. Generation timeout is 60 seconds. `media.syncAnimatedImageToWordAudio` (default `true`) prepends a frozen first frame matching the existing word-audio duration, so the motion starts together with the sentence audio.

### Behavior options

```jsonc
"ankiConnect": {
  "behavior": {
    "overwriteAudio": true,         // replace existing audio, or append
    "overwriteImage": true,         // replace existing image, or append
    "mediaInsertMode": "append",    // "append" or "prepend" to field content
    "autoUpdateNewCards": true,     // auto-update when new card detected
    "highlightWord": true,          // bold the mined word inside the sentence field
    "notificationType": "overlay"   // "overlay", "system", "both", or "none"
  }
}
```

`both` now means overlay + system notification. `osd` and `osd-system` are legacy config-file-only values; set `notificationType` to `"osd-system"` in `config.jsonc` if you previously used `both` and want to keep mpv OSD + system notifications. The Settings window shows `osd` or `osd-system` when already configured, but only offers `overlay`, `system`, `both`, and `none` as normal choices.

When media is available, mined-card overlay and system notifications include the same current-frame thumbnail.

`overwriteAudio` applies to automatic card updates and duplicate-card enrichment. Manual clipboard subtitle updates (`Ctrl/Cmd+C`, then `Ctrl/Cmd+V`) always replace generated sentence audio in `ankiConnect.fields.audio`, even when `overwriteAudio` is disabled.

## Sentence cards (Lapis)

SubMiner can create standalone sentence cards (without a word/expression) using a separate note type. This is designed for use with [Lapis](https://github.com/donkuri/Lapis) and similar sentence-focused note types.

::: warning Required config
Sentence card creation and audio card marking require a non-empty `ankiConnect.isLapis.sentenceCardModel` naming a note type that exists in Anki (default: `"Lapis"`). If the model is empty or missing, the `Ctrl/Cmd+S` and `Ctrl/Cmd+Shift+A` shortcuts will not create cards.
:::

```jsonc
"ankiConnect": {
  "isLapis": {
    "enabled": true,
    "sentenceCardModel": "Lapis" // default; point at your Lapis/Kiku note type
  }
}
```

Trigger with the mine sentence shortcut (`Ctrl/Cmd+S` by default). The card is created directly via AnkiConnect with the sentence, audio, and image filled in.

The dedicated sentence-card and audio-card shortcuts use the Lapis/Kiku-compatible `Sentence` and `SentenceAudio` fields. This does not affect the configured fields used to enrich normal word cards.

To mine multiple subtitle lines as one sentence card, use `Ctrl/Cmd+Shift+S` followed by a digit (1–9) to select how many recent lines to combine.

## Word card type (Kiku/Lapis)

Word cards get a card-type flag when SubMiner fills their sentence, whether that comes from Yomitan auto-enrichment, a manual clipboard update, or stats-dashboard word mining. By default the flag is `IsWordAndSentenceCard`; pick a different one with `ankiConnect.lapisKiku.wordCardKind`.

```jsonc
"ankiConnect": {
  "isKiku": { "enabled": true },
  "lapisKiku": {
    "wordCardKind": "click" // word-and-sentence (default), click, sentence, audio, none
  }
}
```

`click` marks `IsClickCard`, `sentence` marks `IsSentenceCard`, `audio` marks `IsAudioCard`, and `none` leaves the flags untouched for templates that manage them elsewhere. Whichever flag is chosen, the other card-type flags are cleared so the note never claims two card types. The setting is only read when `isKiku` or `isLapis` is enabled, and cards mined with Mine Sentence or Mine Audio keep their own flag.

## Field grouping (Kiku/Senren)

When you mine the same word multiple times, SubMiner can merge the cards instead of creating duplicates. This is designed for note types that support grouped fields: [Kiku](https://github.com/youyoumu/kiku) and [Senren](https://github.com/BrenoAqua/Senren) (which calls the feature scene switching).

```jsonc
"ankiConnect": {
  "isKiku": {
    "enabled": true,
    "fieldGrouping": "manual",         // "auto", "manual", or "disabled"
    "deleteDuplicateInAuto": true      // delete new card after auto-merge
  }
}
```

For Senren note types, enable `isSenren` instead. Kiku and Senren write incompatible markup into the same fields, so only one can be enabled at a time; if both are enabled, Kiku wins and a config warning is emitted.

```jsonc
"ankiConnect": {
  "isSenren": {
    "enabled": true,
    "fieldGrouping": "auto",           // "auto" (default), "manual", or "disabled"
    "deleteDuplicateInAuto": true      // delete new card after auto-merge
  }
}
```

### Modes

**Disabled** (`"disabled"`): No duplicate detection. Each card is independent.

**Auto** (`"auto"`): When a duplicate expression is found, SubMiner merges the new card into the existing one automatically. Both cards' sentences, audio clips, and images are preserved as grouped entries. If `deleteDuplicateInAuto` is true, the new card is deleted after merging.

**Manual** (`"manual"`): A modal appears in the overlay showing both cards. You choose which card to keep, preview the merge result, then confirm. The modal has a 90-second timeout, after which it cancels automatically.

### What gets merged

| Field    | Merge behavior                                  |
| -------- | ----------------------------------------------- |
| Sentence | Both cards' sentences kept as grouped entries   |
| Audio    | Both cards' `[sound:...]` entries kept          |
| Image    | Both cards' images kept                         |
| MiscInfo | Both cards' source info kept as grouped entries |

Identical values from both cards are kept as separate grouped entries; the merge does not deduplicate.

The merge markup depends on the note type. Kiku entries are wrapped in `<span data-group-id="...">` spans ordered newest first. Senren entries follow the [scene switching](https://github.com/BrenoAqua/Senren/blob/main/docs/scene_switching.md) format: sentence, sentenceFurigana, and miscInfo entries use `group` spans when ordinal order is sufficient and numbered `groupN` spans when they need an absolute scene target. Audio and pictures are appended positionally, and the number of sentenceAudio entries drives Senren's scene count. Ungrouped legacy content is wrapped into a group span on first merge, and source `groupN` spans are rebased after the kept note's existing audio scenes.

### Keyboard shortcuts in the modal

| Key         | Action                             |
| ----------- | ---------------------------------- |
| `1` / `2`   | Select card 1 or card 2 to keep    |
| `Enter`     | Confirm selection                  |
| `Backspace` | Go back from the merge preview     |
| `Esc`       | Cancel (keep both cards unchanged) |

## Full config example

```jsonc
{
  "ankiConnect": {
    "enabled": true,
    "url": "http://127.0.0.1:8765",
    "pollingRate": 3000,
    "deck": "",
    "tags": ["SubMiner"],
    "proxy": {
      "enabled": true, // default
      "host": "127.0.0.1",
      "port": 8766,
      "upstreamUrl": "http://127.0.0.1:8765",
    },
    "fields": {
      "word": "Expression",
      "audio": "SentenceAudio",
      "image": "Picture",
      "sentence": "Sentence",
      "miscInfo": "MiscInfo",
    },
    "media": {
      "generateAudio": true,
      "generateImage": true,
      "imageType": "static",
      "imageFormat": "jpg",
      "imageQuality": 92,
      "normalizeAudio": true,
      "mirrorMpvVolume": true,
      "audioPadding": 0,
      "maxMediaDuration": 30,
    },
    "behavior": {
      "overwriteAudio": true,
      "overwriteImage": true,
      "mediaInsertMode": "append",
      "autoUpdateNewCards": true,
      "notificationType": "overlay",
    },
    "metadata": {
      "pattern": "[SubMiner] %f (%t)",
    },
    "isKiku": {
      "enabled": false,
      "fieldGrouping": "disabled",
      "deleteDuplicateInAuto": true,
    },
    "isLapis": {
      "enabled": false,
      "sentenceCardModel": "Lapis",
    },
  },
}
```
