# Mining workflow

SubMiner turns lines from the video you are watching into Anki cards. You look up a word on the overlay, add it with Yomitan, and SubMiner fills in the sentence, an audio clip, and a screenshot. For Anki setup, field mapping, and media options, see [Anki integration](/anki-integration).

## Look up a word

1. Hover the subtitle line on the overlay. Each word is its own hover target.
2. Press your Yomitan scan key or modifier (whatever your Yomitan profile uses, for example `Shift`).
3. The Yomitan popup opens for that word.

Playback pauses while you hover the subtitle and while the Yomitan popup is open. Turn this off with `subtitleStyle.autoPauseVideoOnHover` and `subtitleStyle.autoPauseVideoOnYomitanPopup`.

## Add a word card

Click the add button in the Yomitan popup. SubMiner sees the new note and fills it in:

| Field    | Content                                                |
| -------- | ------------------------------------------------------ |
| Sentence | The current subtitle line, with the mined word in bold |
| Audio    | A clip cut from the video using the subtitle's timing  |
| Image    | A screenshot, or an animated AVIF clip of the line     |
| MiscInfo | Source file name and timestamp                         |

Which note fields receive each item is set in [`ankiConnect.fields`](/anki-integration#field-mapping). With the default proxy mode the card is filled as soon as Yomitan adds it. If you disable the proxy, SubMiner polls Anki and fills the card a few seconds later.

## Update the last card by hand

Use this when auto-update is off, or when the line you want is not the one on screen.

1. Add the word with Yomitan.
2. Press `Ctrl/Cmd+C` to copy the current line. To combine lines, press `Ctrl/Cmd+Shift+C`, then a digit `1` to `9` for how many recent lines to include.
3. Press `Ctrl/Cmd+V`. SubMiner writes the clipboard text into the last-added card's sentence field and adds fresh audio and an image.

A manual update always replaces the sentence audio, even when `ankiConnect.behavior.overwriteAudio` is `false`.

## Mine a sentence card

Press `Ctrl/Cmd+S` to create a sentence card from the current line without a Yomitan lookup. Press `Ctrl/Cmd+Shift+S`, then a digit `1` to `9`, to combine several recent lines into one card. The digit prompt closes after `shortcuts.multiCopyTimeoutMs`.

Sentence cards use the note type named in `ankiConnect.isLapis.sentenceCardModel` and write to its `Sentence` and `SentenceAudio` fields. That note type must exist in Anki. See [sentence cards](/anki-integration#sentence-cards-lapis).

## Mark an audio card

After adding a word, press `Ctrl/Cmd+Shift+A`. SubMiner sets the Lapis/Kiku `IsAudioCard` flag on the last-added card and fills `Sentence`, `SentenceAudio`, the image, and MiscInfo.

## Merge repeated words

If you mine a word you already have a card for, SubMiner can merge the new sentence, audio, and image into the existing card instead of keeping a duplicate. This needs the Kiku or Senren note type with field grouping turned on. In manual mode a dialog shows both cards and lets you pick which one to keep. See [field grouping](/anki-integration#field-grouping-kiku-senren).

## Subtitle display

The overlay has a primary subtitle bar (the Japanese line you mine from) and a secondary bar for a translation track. Each bar is hidden, visible, or shown only on hover.

| Shortcut           | Action                             |
| ------------------ | ---------------------------------- |
| `V`                | Cycle primary bar mode             |
| `Ctrl/Cmd+Shift+V` | Cycle secondary bar mode           |
| `Shift+J`          | Cycle the secondary subtitle track |
| Right-click        | Pause or resume                    |
| Right-click + drag | Move the subtitles                 |

Set the starting modes with `subtitleStyle.primaryDefaultMode` and `secondarySub.defaultMode`. The full list of keys is on [Keyboard shortcuts](/shortcuts).

## Controller

With a gamepad and keyboard-only mode on, you can mine without a mouse: move across words with the left stick, look up with `A`, mine with `X`, and close the popup with `B`. The keyboard keeps working alongside it. See [controller support](/usage#controller-support).

## Fix subtitle timing (subsync)

If the subtitles are out of sync, press `Ctrl+Alt+S` to open the subsync dialog. It uses [alass](https://github.com/kaegi/alass) or [ffsubsync](https://github.com/smacke/ffsubsync), which you install separately.

1. Pick the engine.
2. For alass, pick a reference: a subtitle track with correct timing (the secondary track by default) or the video file.
3. Pick the track to retime (the active primary track by default).
4. Run it. SubMiner loads the retimed subtitle back into the same slot.

For remote streams such as Jellyfin, only alass with a subtitle reference is available, because ffsubsync and the video reference need the local file. If the tools are not found, see [Troubleshooting](/troubleshooting#subtitle-sync-subsync).

## Texthooker

SubMiner can serve a texthooker page at `http://127.0.0.1:5174` that shows each subtitle line as it arrives, so you can do lookups in a browser instead of on the overlay. Start it with `texthooker.launchAtStartup`, the `--texthooker` flag, or the mpv plugin's `texthooker_enabled` option. Change the port with the plugin's `texthooker_port` option. To build your own client, see the [WebSocket / texthooker API](/websocket-texthooker-api).

## Related features

- [Jimaku](/jimaku-integration): search and download subtitle files from the overlay (`Ctrl+Shift+J`).
- [Subtitle annotations](/subtitle-annotations): highlight known words, N+1 targets, JLPT levels, and frequency.
- [Immersion tracking](/immersion-tracking): log watch time and cards mined, and view them in the stats dashboard.
