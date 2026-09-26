# Anki integration

SubMiner talks to Anki through the [AnkiConnect](https://ankiweb.net/shared/info/2055492159) add-on. It fills new cards with the sentence, an audio clip, and a screenshot, and can create sentence cards and merge duplicate words. It is built for the [Lapis](https://github.com/donkuri/lapis), [Kiku](https://kiku.youyoumu.my.id/), and [Senren](https://github.com/BrenoAqua/Senren) note types, but works with any note type once you map its fields.

For the day-to-day flow, see [Mining workflow](/mining-workflow). Every key on this page, with its default, is listed in the [AnkiConnect config reference](/configuration#ankiconnect).

## Prerequisites

1. Install [Anki](https://apps.ankiweb.net/).
2. Install AnkiConnect (add-on code `2055492159`).
3. Install FFmpeg and make sure it is on your `PATH`. SubMiner uses it for audio and images.
4. Keep Anki running while you mine.

If you changed AnkiConnect's port, set `ankiConnect.url` to match.

## How cards get filled

When Yomitan or Hachidori adds a note, SubMiner fills the sentence, audio, image, and MiscInfo fields. It finds new notes in one of two ways:

- **Proxy (default).** SubMiner runs a local AnkiConnect-compatible server. The dictionary sends notes through it, and SubMiner fills each one right after Anki accepts it.
- **Polling.** With `ankiConnect.proxy.enabled` set to `false`, SubMiner asks AnkiConnect for recently added notes every `ankiConnect.pollingRate` milliseconds.

Set `ankiConnect.behavior.autoUpdateNewCards` to `false` to stop automatic filling and update cards by hand with `Ctrl/Cmd+V` instead.

`ankiConnect.deck` limits enrichment and duplicate checks to one deck. If it is empty, SubMiner uses Yomitan's mining deck when it can read it, and otherwise searches all decks.

### Proxy mode setup (Yomitan / texthooker) {#proxy-mode-setup-yomitan-texthooker}

```jsonc
"ankiConnect": {
  "url": "http://127.0.0.1:8765",
  "proxy": {
    "enabled": true,
    "host": "127.0.0.1",
    "port": 8766,
    "upstreamUrl": "http://127.0.0.1:8765"
  }
}
```

Clients must send notes to the proxy (`http://127.0.0.1:8766` here), not to AnkiConnect directly.

- **Bundled Yomitan.** SubMiner sets the active Yomitan profile's Anki server for you. With the proxy on, it always points the profile at the proxy. With the proxy off, it sets `ankiConnect.url`, but only if the profile's server is blank or the stock `http://127.0.0.1:8765`.
- **Browser Yomitan or other clients.** Set the Anki server to the proxy URL yourself. To leave your main profile alone, create a separate Yomitan profile for SubMiner, set its Anki server (Settings, Anki) to the proxy URL, and make it active while you mine.
- **Hachidori.** SubMiner routes Hachidori to the proxy while it is active. Keep the proxy on for screenshots and sentence audio.

### Hachidori settings from SubMiner

With the [Hachidori backend](/usage#hachidori-setup), SubMiner fills Hachidori's first Anki template from your `ankiConnect` settings on startup and whenever you open Hachidori Settings:

- The deck always follows `ankiConnect.deck`, because polling only looks for new cards in that deck.
- Configured tags go into untouched defaults.
- Missing word, sentence, pronunciation-audio, and picture mappings are filled with fields that exist in Anki. Pronunciation uses `fields.wordAudio`, or `fields.audio` when no word-audio field is set.
- If the note type is unset, SubMiner picks the one note type that has your word and sentence fields. Enabled Lapis, Kiku, or Senren narrows the search. A fresh mapping also gets Hachidori's matching preset for readings, definitions, and other known fields.

Apart from the deck, SubMiner only fills missing settings. Custom tags, field mappings, advanced templates, and extra templates stay as you set them. If several note types match, pick one in Hachidori Settings. If Anki was closed, start it and open Hachidori Settings again to retry.

Sentence audio, image timing, translation, metadata, and field grouping stay under SubMiner's control. Pronunciation sources are set in Hachidori. Linking an external dictionary host does not change any of this.

### Proxy troubleshooting

If cards are not getting filled:

1. Check that the proxy is listening while SubMiner runs:

   ```bash
   ss -ltnp | grep 8766
   ```

2. Check that requests pass through to Anki:

   ```bash
   curl -sS http://127.0.0.1:8766 \
     -H 'content-type: application/json' \
     -d '{"action":"version","version":2}'
   ```

3. Read the app log (`app-YYYY-MM-DD.log`) in the logs folder. See [Troubleshooting](/troubleshooting) for where logs live.

## Field mapping

`ankiConnect.fields` maps SubMiner's data to fields on your note type.

| Key                | Receives                                                                  |
| ------------------ | ------------------------------------------------------------------------- |
| `fields.word`      | The mined word                                                            |
| `fields.audio`     | Sentence audio cut from the video                                         |
| `fields.wordAudio` | Read only: Yomitan's word audio, used to time animated images (see below) |
| `fields.image`     | Screenshot or animated clip                                               |
| `fields.sentence`  | Subtitle text                                                             |
| `fields.miscInfo`  | Text from `ankiConnect.metadata.pattern`                                  |

```jsonc
"ankiConnect": {
  "fields": {
    "audio": "SentenceAudio",
    "sentence": "Sentence"
  }
}
```

Field names are matched case-insensitively. A mapped field that is missing from the note type is skipped.

Hachidori prepares downloadable word audio before it saves a note through the proxy, so the animated image delay works on the first mine. Set a downloadable pronunciation source in Hachidori's Audio settings. Browser speech cannot be saved to Anki. Without word audio, Hachidori shows a warning and the card gets no word-audio hold.

`fields.audio` gets sentence audio, not word audio. Yomitan writes its own dictionary audio into your note, so point `fields.audio` at a separate field such as `SentenceAudio`. The default, `ExpressionAudio`, is the field many note types use for Yomitan's word audio, so leaving it would overwrite that audio.

`ankiConnect.tags` adds tags to every mined or updated card. Set it to `[]` to add none.

`ankiConnect.metadata.pattern` builds the MiscInfo text. Tokens: `%f` file name, `%F` file name with extension, `%t` timestamp, `%T` timestamp with milliseconds, `<br>` line break.

## Media

| Key                                  | What it does                                                                                                              |
| ------------------------------------ | ------------------------------------------------------------------------------------------------------------------------- |
| `media.generateAudio`                | Cut sentence audio (MP3) from the subtitle's start and end time                                                           |
| `media.audioPadding`                 | Seconds added before and after the clip                                                                                   |
| `media.fallbackDuration`             | Clip length when the subtitle has no timing                                                                               |
| `media.maxMediaDuration`             | Longest allowed clip, in seconds (`0` removes the cap)                                                                    |
| `media.normalizeAudio`               | Normalize clip loudness                                                                                                   |
| `media.mirrorMpvVolume`              | Scale the clip by mpv's current volume, so quiet playback gives quiet clips                                               |
| `media.generateImage`                | Capture an image                                                                                                          |
| `media.imageType`                    | `static` for one frame, `avif` for an animated clip of the line                                                           |
| `media.imageFormat`                  | Static format: `jpg`, `png`, or `webp`                                                                                    |
| `media.imageQuality`                 | Static image quality                                                                                                      |
| `media.imageMaxWidth` / `Height`     | Static size limit (`0` keeps source size)                                                                                 |
| `media.animatedFps`                  | Animated clip frame rate                                                                                                  |
| `media.animatedMaxWidth` / `Height`  | Animated size limit (`0` keeps aspect ratio)                                                                              |
| `media.animatedCrf`                  | Animated quality, `0` to `63`, lower is better                                                                            |
| `media.syncAnimatedImageToWordAudio` | Hold the first frame for the length of the word audio in `fields.wordAudio`, so the motion starts with the sentence audio |
| `media.reviewTiming`                 | Pause and let you adjust the clip before media is made (see below)                                                        |

Animated AVIF needs an FFmpeg build with an AV1 encoder (`libaom-av1`, `libsvtav1`, or `librav1e`).

Media settings apply to the next card without a restart.

### Review media timing

With `media.reviewTiming` on, SubMiner pauses before making media for word, sentence, and audio cards and opens a review dialog. You can also toggle it for the current session with **Review Media Timing** in the runtime options palette (`Ctrl/Cmd+Shift+O`). Clipboard updates and stats-dashboard mining skip the review.

Playback stays paused while the dialog is open, even if the popup or hover that paused it goes away. When the dialog closes, playback resumes if it was playing before, or if the popup closed in the meantime. A popup that is still open keeps it paused.

The dialog shows the clip over a speech waveform. When the waveform loads, an untouched clip end moves back to just after the last speech in the line. The Line end rail still marks the subtitle's own end.

| Action                   | How                                                                   |
| ------------------------ | --------------------------------------------------------------------- |
| Trim                     | Drag either edge, or click the waveform to move the nearer edge there |
| Nudge an edge            | Arrow keys on a focused edge (100 ms, `Shift` for 500 ms)             |
| Slide the clip           | Drag the middle                                                       |
| Show more timeline       | Earlier / Later                                                       |
| Pick the screenshot      | Screenshot slider or Frame buttons (static images only)               |
| Add previous / next line | `P` / `N` (`Shift+P` / `Shift+N` removes)                             |
| Preview                  | `Space`                                                               |
| Confirm                  | `Enter`                                                               |
| Cancel                   | `Escape`                                                              |

The confirmed range is used as is, with no extra padding. Added lines go into the sentence field. Reset restores the original timing and removes added lines.

When you cancel, you can go back to editing, keep the original timing, create the card without media, or discard it. Discard deletes the Yomitan note or audio card, and skips creation for a sentence card.

### Update behavior

| Key                           | What it does                                         |
| ----------------------------- | ---------------------------------------------------- |
| `behavior.overwriteAudio`     | Replace existing audio instead of adding to it       |
| `behavior.overwriteImage`     | Replace the existing image instead of adding to it   |
| `behavior.mediaInsertMode`    | `append` or `prepend` new media when not overwriting |
| `behavior.autoUpdateNewCards` | Fill new Yomitan notes automatically                 |
| `behavior.highlightWord`      | Bold the mined word in the sentence field            |
| `behavior.notificationType`   | `overlay`, `system`, `both`, or `none`               |

Manual clipboard updates (`Ctrl/Cmd+V`) always replace the sentence audio, whatever `overwriteAudio` says.

## Sentence cards (Lapis) {#sentence-cards-lapis}

`Ctrl/Cmd+S` creates a standalone sentence card from the current line, and `Ctrl/Cmd+Shift+S` then a digit combines several lines. The card uses the note type named in `ankiConnect.isLapis.sentenceCardModel`, which must exist in Anki. If it is empty, no card is created.

```jsonc
"ankiConnect": {
  "isLapis": {
    "enabled": true,
    "sentenceCardModel": "Lapis"
  }
}
```

Sentence cards and audio cards (`Ctrl/Cmd+Shift+A`) always write to the `Sentence` and `SentenceAudio` fields. Normal word cards keep using your `ankiConnect.fields` mapping.

## Word card type (Kiku/Lapis)

When `isKiku` or `isLapis` is enabled, SubMiner sets a card-type flag on word cards it fills. Choose the flag with `ankiConnect.lapisKiku.wordCardKind`:

| Value               | Flag                    |
| ------------------- | ----------------------- |
| `word-and-sentence` | `IsWordAndSentenceCard` |
| `click`             | `IsClickCard`           |
| `sentence`          | `IsSentenceCard`        |
| `audio`             | `IsAudioCard`           |
| `none`              | Leaves flags alone      |

The other card-type flags are cleared. Sentence cards and audio cards keep their own flag.

## Field grouping (Kiku/Senren) {#field-grouping-kiku-senren}

When you mine a word that already has a card, SubMiner can merge the new card into the old one. The sentence, audio, image, and MiscInfo from both cards are kept as grouped entries, and the template lets you switch between them. This works with [Kiku](https://github.com/youyoumu/kiku) and [Senren](https://github.com/BrenoAqua/Senren) (which calls it [scene switching](https://github.com/BrenoAqua/Senren/blob/main/docs/scene_switching.md)).

Enable one of them. They write different markup to the same fields, so only one can be on. If both are enabled, Kiku is used and SubMiner logs a config warning.

```jsonc
"ankiConnect": {
  "isKiku": {
    "enabled": true,
    "fieldGrouping": "manual",
    "deleteDuplicateInAuto": true
  }
}
```

For Senren, use the same keys under `isSenren`.

| `fieldGrouping` | Behavior                                                                        |
| --------------- | ------------------------------------------------------------------------------- |
| `disabled`      | No duplicate check                                                              |
| `auto`          | Merge into the existing card. With `deleteDuplicateInAuto`, delete the new card |
| `manual`        | Show both cards, let you choose which to keep and preview the merge             |

Grouping runs when a new note is added while duplicates exist. With Hachidori, that is the popup's **Add anyway** choice. **Overwrite** updates the existing note in place and only gets media.

The manual dialog cancels itself after 90 seconds. Identical entries are not deduplicated. Press `Ctrl/Cmd+G` to run the duplicate check on the last card yourself.

| Key         | Action                                |
| ----------- | ------------------------------------- |
| `1` / `2`   | Keep card 1 or card 2                 |
| `Enter`     | Confirm                               |
| `Backspace` | Back from the merge preview           |
| `Esc`       | Cancel and leave both cards unchanged |

## Config validation

Invalid `ankiConnect` values produce a warning and fall back to the default. Use JSON booleans (`true`, not `"true"`) and a positive number for `pollingRate`.
