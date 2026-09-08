# Mining workflow

This guide walks the whole sentence mining loop, from starting a video to ending up with an Anki card that has audio, a screenshot, and the surrounding sentence.

## Overview

_Sentence mining_ means turning sentences you hit while watching native video into Anki cards, so you learn a word in the context where you first met it. The idea is old. The tedious part is everything between spotting the word and having a finished card, and that is the part SubMiner does for you.

SubMiner draws a transparent overlay on top of mpv and renders each subtitle line as interactive text. Hover a word, trigger a Yomitan lookup with your configured key or modifier, then add the card. SubMiner attaches the sentence, an audio clip, and a screenshot on its own, so there is nothing to copy-paste or screenshot by hand.

> **Yomitan** is the popup dictionary that shows definitions when you hover or scan a word. **AnkiConnect** is the add-on that lets SubMiner talk to Anki. Both are set up during installation - see [Anki Integration](/anki-integration) if you have not configured them yet.

## Creating Anki cards

There are four ways to create or enrich cards, depending on your workflow.

### 1. Auto-update from Yomitan

This is the most common flow. Yomitan creates a card in Anki, and SubMiner enriches it automatically.

1. Hover a word, then trigger Yomitan lookup → Yomitan popup appears.
2. Click the Anki icon in Yomitan to add the word.
3. SubMiner receives or detects the new card:
   - **Proxy mode** (default, `ankiConnect.proxy.enabled: true`): immediate enrich after a successful `addNote` / `addNotes` is pushed through the local proxy.
   - **Polling mode** (fallback, when the proxy is disabled): detects new cards via AnkiConnect polling (`ankiConnect.pollingRate`, default 3 seconds).
4. SubMiner updates the card with:
   - **Sentence**: The current subtitle line.
   - **Audio**: Extracted from the video using the subtitle's start/end timing (plus optional configured padding).
   - **Image**: A screenshot or animated clip from the current playback position.
   - **MiscInfo**: Metadata like filename and timestamp.

Configure which fields to fill in `ankiConnect.fields`. See [Anki Integration](/anki-integration) for details.

### 2. manual update from clipboard

If you prefer a hands-on approach (animecards-style), you can copy the current subtitle to the clipboard and then paste it onto the last-added Anki card:

1. Add a word via Yomitan as usual.
2. Press `Ctrl/Cmd+C` to copy the current subtitle line to the clipboard.
   - For multiple lines: press `Ctrl/Cmd+Shift+C`, then a digit `1`–`9` to select how many recent subtitle lines to combine. The combined text is copied to the clipboard.
3. Press `Ctrl/Cmd+V` to update the last-added card with the clipboard contents plus audio and image, the same fields auto-update would fill.

Manual clipboard updates always replace generated sentence audio in `ankiConnect.fields.audio`, even when `ankiConnect.behavior.overwriteAudio` is disabled. Normal word-card updates use the configured sentence and audio fields even when Lapis or Kiku support is enabled.

Use this when auto-update is off, or when the line you want on the card is not the line currently on screen.

| Shortcut                   | Action                          | Config key                              |
| -------------------------- | ------------------------------- | --------------------------------------- |
| `Ctrl/Cmd+C`               | Copy current subtitle           | `shortcuts.copySubtitle`                |
| `Ctrl/Cmd+Shift+C` + digit | Copy multiple recent lines      | `shortcuts.copySubtitleMultiple`        |
| `Ctrl/Cmd+V`               | Update last card from clipboard | `shortcuts.updateLastCardFromClipboard` |

### 3. mine Sentence (hotkey)

Create a standalone sentence card without going through Yomitan:

- **Mine current sentence**: `Ctrl/Cmd+S` (configurable via `shortcuts.mineSentence`)
- **Mine multiple lines**: `Ctrl/Cmd+Shift+S` followed by a digit 1–9 to select how many recent subtitle lines to combine (the digit selector times out after 3 seconds, configurable via `shortcuts.multiCopyTimeoutMs`).

The sentence card uses the note type configured in `isLapis.sentenceCardModel` and always maps sentence/audio to `Sentence` and `SentenceAudio`.

::: warning Requires Lapis/Kiku note type
Sentence card creation requires `ankiConnect.isLapis.sentenceCardModel` to name a [Lapis](https://github.com/donkuri/lapis) or [Kiku](https://github.com/youyoumu/kiku) compatible note type that exists in Anki (default: `"Lapis"`). See [Anki Integration - Sentence Cards](/anki-integration#sentence-cards-lapis) for setup.
:::

### 4. mark as audio card

After adding a word via Yomitan, press the audio card shortcut (`Ctrl/Cmd+Shift+A` by default, `shortcuts.markAudioCard`) to mark the card as an audio card. This sets the audio-card flag and fills sentence, image, and metadata fields alongside the full-subtitle audio clip.

::: warning Requires Lapis/Kiku note type
Audio card marking uses the same `ankiConnect.isLapis.sentenceCardModel` note type as sentence cards. See [Anki Integration - Sentence Cards](/anki-integration#sentence-cards-lapis) for setup.
:::

### Field grouping (Kiku/Senren)

If you mine the same word from different sentences, SubMiner can merge the cards instead of creating duplicates. This is built for [Kiku](https://github.com/youyoumu/kiku) and [Senren](https://github.com/BrenoAqua/Senren) note types that support grouped fields (Senren calls it scene switching).

1. You add a word via Yomitan.
2. SubMiner detects the new card and checks if a card with the same expression already exists.
3. If a duplicate is found (this requires Kiku or Senren to be enabled with a field grouping mode of `"auto"` or `"manual"`):
   - **Auto mode**: Merges automatically. Both sentences, audio clips, images, and source info are combined into the existing card. The duplicate is optionally deleted.
   - **Manual mode**: A modal appears showing both cards side by side. You choose which card to keep and preview the merged result before confirming.

See [Anki Integration - Field Grouping](/anki-integration#field-grouping-kiku-senren) for configuration options, merge behavior, and modal keyboard shortcuts.

## Overlay model

SubMiner uses one overlay window with modal surfaces. It carries two subtitle bars - a primary reading bar and a secondary translation/context bar - plus modal dialogs that open on top.

Toggle the entire overlay window with `Alt+Shift+O` (global) or `y-t` (mpv plugin).

### Primary subtitle layer

The primary bar renders each subtitle as separate hoverable word spans, each carrying its reading and headword. Its styling is independent of mpv's own subtitle rendering. It supports:

- Word-level hover targets for Yomitan lookup
- Auto pause/resume on subtitle hover (enabled by default via `subtitleStyle.autoPauseVideoOnHover`)
- Auto pause/resume while the Yomitan popup is open (enabled by default via `subtitleStyle.autoPauseVideoOnYomitanPopup`)
- Right-click to pause/resume
- Right-click + drag to reposition subtitles
- **Reading annotations** - known words, N+1 targets, character-name matches, JLPT levels, and frequency hits can all be visually highlighted

### Secondary subtitle bar

The secondary bar is a compact top-strip region in the same overlay window. It shows a secondary subtitle track, usually English, above the primary reading line. Use it to sanity-check your comprehension without breaking out of the mining flow.

For local media, SubMiner can parse supported embedded secondary tracks into timed cues. For remote URLs and files on network mounts, it uses mpv's live secondary subtitle text instead of scanning the media with ffmpeg.

The `secondarySub` config controls it, and it opens and closes with the main overlay window. Cycle which track feeds it with `Shift+J`.

SubMiner collapses duplicate ASS layers in parsed secondary tracks. Exact repeated lines collapse at any length, while distinct simultaneous short lines remain separate. Long dialogue and positioned-sign copies also collapse when they differ only in whitespace or terminal punctuation. Dense multi-row sign layouts, such as translated timetables, are excluded instead of being concatenated into the secondary bar.

### Display modes

Both the primary and secondary subtitle bars share the same three visibility modes, and each can be changed independently at runtime:

- **Hidden** - the bar is not shown.
- **Visible** - the bar is always shown.
- **Hover** - the bar is revealed only while you hover over the overlay.

By default the **primary** bar is `visible` (`subtitleStyle.primaryDefaultMode`) and the **secondary** bar is `hover` (`secondarySub.defaultMode`).

Cycle each bar's mode at runtime with its own shortcut:

| Shortcut           | Action                                                   | Config key                     |
| ------------------ | -------------------------------------------------------- | ------------------------------ |
| `V`                | Cycle primary subtitle mode (hidden → visible → hover)   | overlay-local                  |
| `Ctrl/Cmd+Shift+V` | Cycle secondary subtitle mode (hidden → visible → hover) | `shortcuts.toggleSecondarySub` |

### Modal surfaces

Jimaku search, field-grouping, runtime options, and manual subsync open as modal surfaces on top of the same overlay window.

## Looking up words

1. Hover over the subtitle area - the overlay activates pointer events.
2. Hover the word you want. SubMiner keeps per-token boundaries so Yomitan can target that token cleanly.
3. Trigger Yomitan lookup with your configured lookup key/modifier (for example `Shift` if that is how your Yomitan profile is set up).
4. Yomitan opens its lookup popup for the hovered token.
5. From the popup, add the word to Anki.

### Controller workflow

With a gamepad connected and keyboard-only mode enabled, the full mining loop works without a mouse or keyboard:

1. **Navigate** - push the left stick left/right to move the token highlight across subtitle words.
2. **Look up** - press `A` to trigger Yomitan lookup on the highlighted word.
3. **Browse the popup** - push the left stick up/down to smooth-scroll through the Yomitan popup, or use the right stick for larger jumps.
4. **Cycle audio** - press `R1` to move to the next dictionary audio entry, `L1` to play the current one.
5. **Mine** - press `X` to create an Anki card for the current sentence (same as `Ctrl+S`).
6. **Close** - press `B` to dismiss the Yomitan popup and return to subtitle navigation.
7. **Pause/resume** - press `L3` (left stick click) to toggle mpv pause at any time.

Once controller support is on, the controller and keyboard both stay live. You can drop the controller mid-episode and keep going with the keyboard. Toggle keyboard-only mode with `Y` on the controller.

See [Usage - Controller Support](/usage#controller-support) for setup details and [Configuration - Controller Support](/configuration#controller-support) for the full mapping and tuning options.

## Subtitle sync (subsync)

If your subtitle file is out of sync with the audio, SubMiner can resynchronize it using [alass](https://github.com/kaegi/alass) or [ffsubsync](https://github.com/smacke/ffsubsync).

1. Open the subsync modal from the overlay.
2. Select the sync engine (alass or ffsubsync).
3. For alass, pick the **reference** - the subtitle with correct timing. This defaults to the secondary subtitle track. The loaded video file can also be used as the reference (alass extracts the audio itself), but it is never the default.
4. Pick the **out-of-sync subtitle** - the track that gets retimed. This defaults to the active primary subtitle track and applies to both engines.
5. SubMiner runs the sync and reloads the corrected subtitle into the slot the out-of-sync track came from: retiming the secondary track keeps it secondary and leaves the primary track selected.

The reference and the out-of-sync subtitle must be different tracks; the reference list hides whichever track is selected as the target.

For remote streams, including Jellyfin playback, the modal only offers alass with a subtitle reference. Jellyfin subtitle URLs are cached as temporary subtitle files so alass can read them, but the video stream is not downloaded. ffsubsync and the video-file reference need direct access to the local media file and are unavailable for stream URLs.

Install the sync tools separately - see [Troubleshooting](/troubleshooting#subtitle-sync-subsync) if the tools are not found.

## Texthooker

SubMiner serves a texthooker UI from a local HTTP server at `http://127.0.0.1:5174`. The port is fixed unless you override it with the mpv plugin's `texthooker_port` script-opt. External tools read subtitle text from it as lines arrive, which is how you would feed a browser-based Yomitan instance.

The texthooker page displays the current subtitle and updates as new lines arrive. This is useful if you prefer to do lookups in a browser rather than through the overlay's built-in Yomitan.

If you want to build your own browser client, websocket consumer, or automation relay, see [WebSocket / Texthooker API & Integration](/websocket-texthooker-api).

## Related features

These feed into the mining loop but each has its own page:

- **[Jimaku subtitle search](/jimaku-integration)** - search and download anime subtitle files directly from the overlay (`Ctrl+Shift+J` by default), then load them into mpv.
- **[N+1 word highlighting](/subtitle-annotations#n-1-word-highlighting)** - reads your Anki decks and highlights words you already know, so a line with exactly one unknown word stands out while you watch.
- **[Immersion tracking](/immersion-tracking)** - log watching and mining activity to a local database and view session times, words seen, and cards mined in the built-in stats dashboard.

Next: [Anki Integration](/anki-integration) - field mapping, media generation, and card enrichment configuration.
