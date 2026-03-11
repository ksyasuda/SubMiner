# Anki Integration

SubMiner uses the [AnkiConnect](https://ankiweb.net/shared/info/2055492159) add-on to create and update Anki cards with sentence context, audio, and screenshots.
This project is built primarily for [Kiku](https://kiku.youyoumu.my.id/) and [Lapis](https://github.com/donkuri/lapis) note types, including sentence-card and field-grouping behavior.

## Prerequisites

1. Install [Anki](https://apps.ankiweb.net/).
2. Install the [AnkiConnect](https://ankiweb.net/shared/info/2055492159) add-on (code: `2055492159`).
3. Keep Anki running while using SubMiner.

AnkiConnect listens on `http://127.0.0.1:8765` by default. If you changed the port in AnkiConnect's settings, update `ankiConnect.url` in your SubMiner config.

## Auto-Enrichment Transport

When you add a word via Yomitan, SubMiner detects the new card and fills in the sentence, audio, image, and translation fields automatically. Two detection methods are available:

**Proxy mode** — SubMiner runs a local AnkiConnect-compatible proxy and intercepts card creation instantly. Recommended when possible.

**Polling mode** (default) — SubMiner polls AnkiConnect every few seconds for newly added cards. Simpler setup, but with a short delay (~3 seconds).

Use proxy mode if you want immediate enrichment. Use polling mode if your Yomitan instance is external (browser-based) or you prefer minimal configuration.

In both modes, the enrichment workflow is the same:

1. Checks if a duplicate expression already exists (for field grouping).
2. Updates the sentence field with the current subtitle.
3. Generates and uploads audio and image media.
4. Fills the translation field from the secondary subtitle or AI.
5. Writes metadata to the miscInfo field.

Polling mode uses the query `"deck:<your-deck>" added:1` to find recently added cards. If no deck is configured, it searches all decks.

### Proxy Mode Setup (Yomitan / Texthooker)

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

When SubMiner loads the bundled Yomitan extension, it also attempts to update the **default Yomitan profile** (`profiles[0].options.anki.server`) to the active SubMiner endpoint:

- proxy URL when `ankiConnect.proxy.enabled` is `true`
- direct `ankiConnect.url` when proxy mode is disabled

To avoid clobbering custom setups, this auto-update only changes the default profile when its current server is blank or the stock Yomitan default (`http://127.0.0.1:8765`).

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

This is only for non-bundled, external/browser Yomitan or other clients. The bundled profile auto-update logic only targets `profiles[0]` when it is blank or still default.

### Proxy Troubleshooting (quick checks)

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

3. Check both log sinks:

- Launcher/mpv-integrated log: `~/.cache/SubMiner/mp.log`
- App runtime log: `~/.config/SubMiner/logs/SubMiner-YYYY-MM-DD.log`

4. Ensure config JSONC is valid and logging shape is correct:

```jsonc
"logging": {
  "level": "debug"
}
```

`"logging": "debug"` is invalid for current schema and can break reload/start behavior.

## Field Mapping

SubMiner maps its data to your Anki note fields. Configure these under `ankiConnect.fields`:

```jsonc
"ankiConnect": {
  "fields": {
    "audio": "ExpressionAudio",    // audio clip from the video
    "image": "Picture",             // screenshot or animated clip
    "sentence": "Sentence",         // subtitle text
    "miscInfo": "MiscInfo",         // metadata (filename, timestamp)
    "translation": "SelectionText"  // secondary sub or AI translation
  }
}
```

Field names must match your Anki note type exactly (case-sensitive). If a configured field does not exist on the note type, SubMiner skips it without error.

### Minimal Config

If you only want sentence and audio on your cards:

```jsonc
"ankiConnect": {
  "enabled": true,
  "fields": {
    "sentence": "Sentence",
    "audio": "ExpressionAudio"
  }
}
```

## Media Generation

SubMiner uses FFmpeg to generate audio and image media from the video. FFmpeg must be installed and on `PATH`.

### Audio

Audio is extracted from the video file using the subtitle's start and end timestamps, with configurable padding added before and after.

```jsonc
"ankiConnect": {
  "media": {
    "generateAudio": true,
    "audioPadding": 0.5,         // seconds before and after subtitle timing
    "maxMediaDuration": 30       // cap total duration in seconds
  }
}
```

Output format: MP3 at 44100 Hz. If the video has multiple audio streams, SubMiner uses the active stream.

The audio is uploaded to Anki's media folder and inserted as `[sound:audio_<timestamp>.mp3]`.

### Screenshots (Static)

A single frame is captured at the current playback position.

```jsonc
"ankiConnect": {
  "media": {
    "generateImage": true,
    "imageType": "static",
    "imageFormat": "jpg",        // "jpg", "png", or "webp"
    "imageQuality": 92,          // 1–100
    "imageMaxWidth": null,       // optional, preserves aspect ratio
    "imageMaxHeight": null
  }
}
```

### Animated Clips (AVIF)

Instead of a static screenshot, SubMiner can generate an animated AVIF covering the subtitle duration.

```jsonc
"ankiConnect": {
  "media": {
    "generateImage": true,
    "imageType": "avif",
    "animatedFps": 10,
    "animatedMaxWidth": 640,
    "animatedMaxHeight": null,
    "animatedCrf": 35            // 0–63, lower = better quality
  }
}
```

Animated AVIF requires an AV1 encoder (`libaom-av1`, `libsvtav1`, or `librav1e`) in your FFmpeg build. Generation timeout is 60 seconds.

### Behavior Options

```jsonc
"ankiConnect": {
  "behavior": {
    "overwriteAudio": true,         // replace existing audio, or append
    "overwriteImage": true,         // replace existing image, or append
    "mediaInsertMode": "append",    // "append" or "prepend" to field content
    "autoUpdateNewCards": true,     // auto-update when new card detected
    "notificationType": "osd"       // "osd", "system", "both", or "none"
  }
}
```

## AI Translation

SubMiner can auto-translate the mined sentence and fill the translation field.
Secondary subtitle text still wins when present. AI translation is only attempted when `ankiConnect.ai.enabled` is `true` and no secondary subtitle exists.

```jsonc
"ai": {
  "enabled": true,
  "apiKey": "sk-...",
  "apiKeyCommand": "",
  "baseUrl": "https://openrouter.ai/api",
  "requestTimeoutMs": 15000
},
"ankiConnect": {
  "ai": {
    "enabled": true,
    "model": "openai/gpt-4o-mini",
    "systemPrompt": "Translate mined sentence text only."
  }
}
```

`ankiConnect.ai` controls feature-local enablement plus optional `model` / `systemPrompt` overrides.
Provider credentials and request transport settings live in top-level `ai`.

Translation priority:

1. If a secondary subtitle is available, use it as the translation.
2. If `ankiConnect.ai.enabled` is `true` and top-level `ai.enabled` is `true`, call the shared AI provider.
3. If AI translation fails and no secondary subtitle exists, fall back to the original sentence text.

The built-in translation request asks for English output by default. Customize that behavior through `ankiConnect.ai.systemPrompt`.

## Sentence Cards (Lapis)

SubMiner can create standalone sentence cards (without a word/expression) using a separate note type. This is designed for use with [Lapis](https://github.com/donkuri/Lapis) and similar sentence-focused note types.

```jsonc
"ankiConnect": {
  "isLapis": {
    "enabled": true,
    "sentenceCardModel": "Japanese sentences"
  }
}
```

Trigger with the mine sentence shortcut (`Ctrl/Cmd+S` by default). The card is created directly via AnkiConnect with the sentence, audio, and image filled in.

To mine multiple subtitle lines as one sentence card, use `Ctrl/Cmd+Shift+S` followed by a digit (1–9) to select how many recent lines to combine.

## Field Grouping (Kiku)

When you mine the same word multiple times, SubMiner can merge the cards instead of creating duplicates. This is designed for note types like [Kiku](https://github.com/youyoumu/kiku) that support grouped sentence/audio/image fields.

```jsonc
"ankiConnect": {
  "isKiku": {
    "enabled": true,
    "fieldGrouping": "manual",         // "auto", "manual", or "disabled"
    "deleteDuplicateInAuto": true      // delete new card after auto-merge
  }
}
```

### Modes

**Disabled** (`"disabled"`): No duplicate detection. Each card is independent.

**Auto** (`"auto"`): When a duplicate expression is found, SubMiner merges the new card into the existing one automatically. Both sentences, audio clips, and images are preserved, and exact duplicate values are collapsed to one entry. If `deleteDuplicateInAuto` is true, the new card is deleted after merging.

**Manual** (`"manual"`): A modal appears in the overlay showing both cards. You choose which card to keep, preview the merge result, then confirm. The modal has a 90-second timeout, after which it cancels automatically.

### What Gets Merged

| Field    | Merge behavior                                                  |
| -------- | --------------------------------------------------------------- |
| Sentence | Both sentences preserved (exact duplicate text is deduplicated) |
| Audio    | Both `[sound:...]` entries kept (exact duplicates deduplicated) |
| Image    | Both images kept (exact duplicates deduplicated)                |

### Keyboard Shortcuts in the Modal

| Key       | Action                             |
| --------- | ---------------------------------- |
| `1` / `2` | Select card 1 or card 2 to keep    |
| `Enter`   | Confirm selection                  |
| `Esc`     | Cancel (keep both cards unchanged) |

## Full Config Example

```jsonc
{
  "ankiConnect": {
    "enabled": true,
    "url": "http://127.0.0.1:8765",
    "pollingRate": 3000,
    "proxy": {
      "enabled": false,
      "host": "127.0.0.1",
      "port": 8766,
      "upstreamUrl": "http://127.0.0.1:8765",
    },
    "fields": {
      "audio": "ExpressionAudio",
      "image": "Picture",
      "sentence": "Sentence",
      "miscInfo": "MiscInfo",
      "translation": "SelectionText",
    },
    "media": {
      "generateAudio": true,
      "generateImage": true,
      "imageType": "static",
      "imageFormat": "jpg",
      "imageQuality": 92,
      "audioPadding": 0.5,
      "maxMediaDuration": 30,
    },
    "behavior": {
      "overwriteAudio": true,
      "overwriteImage": true,
      "mediaInsertMode": "append",
      "autoUpdateNewCards": true,
      "notificationType": "osd",
    },
    "ai": {
      "enabled": false,
      "model": "openai/gpt-4o-mini",
      "systemPrompt": "Translate mined sentence text only.",
    },
    "isKiku": {
      "enabled": false,
      "fieldGrouping": "disabled",
      "deleteDuplicateInAuto": true,
    },
    "isLapis": {
      "enabled": false,
      "sentenceCardModel": "Japanese sentences",
    },
  },
  "ai": {
    "enabled": false,
    "apiKey": "",
    "apiKeyCommand": "",
    "baseUrl": "https://openrouter.ai/api",
    "requestTimeoutMs": 15000,
  },
}
```
