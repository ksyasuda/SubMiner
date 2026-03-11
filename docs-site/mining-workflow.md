# Mining Workflow

This guide walks through the sentence mining loop — from watching a video to creating Anki cards with audio, screenshots, and context.

## Overview

SubMiner runs as a transparent overlay on top of mpv. As subtitles play, the overlay displays them as interactive text. You click a word to look it up with Yomitan, then create an Anki card with a single action. SubMiner automatically attaches the sentence, audio clip, and screenshot.

```text
Watch video → See subtitle → Click word → Yomitan lookup → Add to Anki
                                                              ↓
                                              SubMiner auto-fills:
                                              sentence, audio, image, translation
```

## Subtitle Delivery Path (Startup + Runtime)

SubMiner prioritizes subtitle responsiveness over heavy initialization:

1. The first subtitle render is **plain text first** (no tokenization wait).
2. Tokenized enrichment (word spans, known-word flags, JLPT/frequency metadata) is applied right after parsing completes.
3. Under rapid subtitle churn, SubMiner uses a **latest-only tokenization queue** so stale lines are dropped instead of building lag.
4. MeCab, Yomitan extension load, and dictionary prewarm run as background warmups after overlay initialization (configurable via `startupWarmups`, including low-power mode).

This keeps early playback snappy and avoids mpv-side sluggishness while startup work completes.

## Overlay Model

SubMiner uses one overlay window with modal surfaces.

### Primary Subtitle Layer

The visible overlay renders subtitles as tokenized, clickable word spans. Each word is a separate element with reading and headword data attached. This plane is styled independently from mpv subtitles and supports:

- Word-level click targets for Yomitan lookup
- Auto pause/resume on subtitle hover (enabled by default via `subtitleStyle.autoPauseVideoOnHover`)
- Optional pause while the Yomitan popup is open (`subtitleStyle.autoPauseVideoOnYomitanPopup`)
- Right-click to pause/resume
- Right-click + drag to reposition subtitles
- Modal dialogs for Jimaku search, field grouping, subsync, and runtime options
- **Reading annotations** — known words, N+1 targets, character-name matches, JLPT levels, and frequency hits can all be visually highlighted

Toggle visibility with `Alt+Shift+O` (global) or `y-t` (mpv plugin).

### Secondary Subtitle Bar

The secondary subtitle bar is a compact top-strip region in the same overlay window for translation/context visibility while keeping primary reading flow below. It mirrors your configured secondary subtitle preference and can be independently shown or hidden.

It is controlled by `secondarySub` configuration and shares lifecycle with the main overlay window.

### Modal Surfaces

Jimaku search, field-grouping, runtime options, and manual subsync open as modal surfaces on top of the same overlay window.

## Looking Up Words

1. Hover over the subtitle area — the overlay activates pointer events.
2. Click any word. SubMiner uses Unicode-aware boundary detection (`Intl.Segmenter`) to select it. On macOS, hovering is enough.
3. Yomitan detects the selection and opens its lookup popup.
4. From the popup, add the word to Anki.

## Creating Anki Cards

There are three ways to create cards, depending on your workflow.

### 1. Auto-Update from Yomitan

This is the most common flow. Yomitan creates a card in Anki, and SubMiner enriches it automatically.

1. Click a word → Yomitan popup appears.
2. Click the Anki icon in Yomitan to add the word.
3. SubMiner receives or detects the new card:
   - **Proxy mode** (`ankiConnect.proxy.enabled: true`): immediate enrich after successful `addNote` / `addNotes`.
   - **Polling mode** (default): detects via AnkiConnect polling (`ankiConnect.pollingRate`, default 3 seconds).
4. SubMiner updates the card with:
   - **Sentence**: The current subtitle line.
   - **Audio**: Extracted from the video using the subtitle's start/end timing (plus configurable padding).
   - **Image**: A screenshot or animated clip from the current playback position.
   - **Translation**: From the secondary subtitle track, or generated via AI if configured.
   - **MiscInfo**: Metadata like filename and timestamp.

Configure which fields to fill in `ankiConnect.fields`. See [Anki Integration](/anki-integration) for details.

### 2. Manual Update from Clipboard

If you prefer a hands-on approach (animecards-style), you can copy the current subtitle to the clipboard and then paste it onto the last-added Anki card:

1. Add a word via Yomitan as usual.
2. Press `Ctrl/Cmd+C` to copy the current subtitle line to the clipboard.
   - For multiple lines: press `Ctrl/Cmd+Shift+C`, then a digit `1`–`9` to select how many recent subtitle lines to combine. The combined text is copied to the clipboard.
3. Press `Ctrl/Cmd+V` to update the last-added card with the clipboard contents plus audio, image, and translation — the same fields auto-update would fill.

This is useful when auto-update is disabled or when you want explicit control over which subtitle line gets attached to the card.

| Shortcut                   | Action                          | Config key                              |
| -------------------------- | ------------------------------- | --------------------------------------- |
| `Ctrl/Cmd+C`               | Copy current subtitle           | `shortcuts.copySubtitle`                |
| `Ctrl/Cmd+Shift+C` + digit | Copy multiple recent lines      | `shortcuts.copySubtitleMultiple`        |
| `Ctrl/Cmd+V`               | Update last card from clipboard | `shortcuts.updateLastCardFromClipboard` |

### 3. Mine Sentence (Hotkey)

Create a standalone sentence card without going through Yomitan:

- **Mine current sentence**: `Ctrl/Cmd+S` (configurable via `shortcuts.mineSentence`)
- **Mine multiple lines**: `Ctrl/Cmd+Shift+S` followed by a digit 1–9 to select how many recent subtitle lines to combine.

The sentence card uses the note type configured in `isLapis.sentenceCardModel` and always maps sentence/audio to `Sentence` and `SentenceAudio`.

### 4. Mark as Audio Card

After adding a word via Yomitan, press the audio card shortcut to overwrite the audio with a longer clip spanning the full subtitle timing.

## Secondary Subtitles

SubMiner can display a secondary subtitle track (typically English) alongside the primary Japanese subtitles. This is useful for:

- Quick comprehension checks without leaving the mining flow.
- Auto-populating the translation field on mined cards.

### Display Modes

Cycle through modes with the configured shortcut:

- **Hidden**: Secondary subtitle not shown.
- **Visible**: Always displayed below the primary subtitle.
- **Hover**: Only shown when you hover over the primary subtitle.

When a card is created, SubMiner uses the secondary subtitle text as the translation field value (unless AI translation is configured to override it).

## Field Grouping (Kiku)

If you mine the same word from different sentences, SubMiner can merge the cards instead of creating duplicates. This feature is designed for use with [Kiku](https://github.com/youyoumu/kiku) and similar note types that support grouped fields.

### How It Works

1. You add a word via Yomitan.
2. SubMiner detects the new card and checks if a card with the same expression already exists.
3. If a duplicate is found:
   - **Auto mode** (`fieldGrouping: "auto"`): Merges automatically. Both sentences, audio clips, and images are combined into the existing card. The duplicate is optionally deleted.
   - **Manual mode** (`fieldGrouping: "manual"`): A modal appears showing both cards side by side. You choose which card to keep and preview the merged result before confirming.

See [Anki Integration — Field Grouping](/anki-integration#field-grouping-kiku) for configuration options, merge behavior, and modal keyboard shortcuts.

## Jimaku Subtitle Search

SubMiner integrates with [Jimaku](https://jimaku.cc) to search and download subtitle files for anime directly from the overlay.

1. Open the Jimaku modal via the configured shortcut (`Ctrl+Shift+J` by default).
2. SubMiner auto-fills the search from the current video filename (title, season, episode).
3. Browse matching entries and select a subtitle file to download.
4. The subtitle is loaded into mpv as a new track.

Requires an internet connection. An API key (`jimaku.apiKey`) is optional but recommended for higher rate limits.

## Texthooker

SubMiner runs a local HTTP server at `http://127.0.0.1:5174` (configurable port) that serves a texthooker UI. This allows external tools — such as a browser-based Yomitan instance — to receive subtitle text in real time.

The texthooker page displays the current subtitle and updates as new lines arrive. This is useful if you prefer to do lookups in a browser rather than through the overlay's built-in Yomitan.

## Subtitle Sync (Subsync)

If your subtitle file is out of sync with the audio, SubMiner can resynchronize it using [alass](https://github.com/kaegi/alass) or [ffsubsync](https://github.com/smacke/ffsubsync).

1. Open the subsync modal from the overlay.
2. Select the sync engine (alass or ffsubsync).
3. For alass, select a reference subtitle track from the video.
4. SubMiner runs the sync and reloads the corrected subtitle.

Install the sync tools separately — see [Troubleshooting](/troubleshooting#subtitle-sync-subsync) if the tools are not found.

## N+1 Word Highlighting

When enabled, SubMiner cross-references your Anki decks to highlight known words in the overlay, making true N+1 sentences (exactly one unknown word) easy to spot during immersion.

See [Subtitle Annotations — N+1](/subtitle-annotations#n1-word-highlighting) for configuration options and color settings.

## Immersion Tracking

SubMiner can log your watching and mining activity to a local SQLite database — session times, words seen, cards mined, and daily/monthly rollups.

Enable it in your config:

```jsonc
"immersionTracking": {
  "enabled": true,
  "dbPath": ""  // leave empty to use the default location
}
```

See [Immersion Tracking](/immersion-tracking) for the full schema and retention settings.

Next: [Anki Integration](/anki-integration) — field mapping, media generation, and card enrichment configuration.
