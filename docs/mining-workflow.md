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

SubMiner now prioritizes subtitle responsiveness over heavy initialization:

1. The first subtitle render is **plain text first** (no tokenization wait).
2. Tokenized enrichment (word spans, known-word flags, JLPT/frequency metadata) is applied right after parsing completes.
3. Under rapid subtitle churn, SubMiner uses a **latest-only tokenization queue** so stale lines are dropped instead of building lag.
4. MeCab, Yomitan extension load, and dictionary prewarm run as background warmups after overlay initialization.

This keeps early playback snappy and avoids mpv-side sluggishness while startup work completes.

## The Three Overlay Planes

SubMiner uses three overlay planes, each serving a different purpose.

### Visible Overlay

The visible overlay renders subtitles as tokenized, clickable word spans. Each word is a separate element with reading and headword data attached. This plane is styled independently from mpv subtitles and supports:

- Word-level click targets for Yomitan lookup
- Right-click to pause/resume
- Right-click + drag to reposition subtitles
- Modal dialogs for Jimaku search, field grouping, subsync, and runtime options
- **N+1 highlighting** — known words from your Anki deck are visually highlighted

Toggle with `Alt+Shift+O` (global) or `y-t` (mpv plugin).

### Secondary Subtitle Plane

The secondary plane is a compact top-strip layer for translation and context visibility while keeping primary reading flow below. It mirrors your configured secondary subtitle preference and can be independently shown or hidden.

It is controlled by `secondarySub` configuration and shares lifecycle with the overlay stack.

### Invisible Overlay

The invisible overlay is a transparent layer aligned with mpv's own subtitle rendering. It uses mpv's subtitle metrics (font size, margins, position, scaling) to map click targets accurately.

This layer still supports:

- Word-level click-through lookups over the text region
- Optional manual position fine-tuning in pixel mode
- Independent toggle behavior with global shortcuts

Position edit mode is available via `Ctrl/Cmd+Shift+P`, then arrow keys / `hjkl` to nudge position; `Shift` moves faster. Save with `Enter` or `Ctrl+S`, cancel with `Esc`.

Toggle controls:

- `Alt+Shift+O` / `y-t`: visible overlay
- `Alt+Shift+I` / `y-i`: invisible overlay
- Secondary plane visibility is controlled via `secondarySub` config and matching global shortcuts.

## Looking Up Words

### On the Visible Overlay

1. Hover over the subtitle area — the overlay activates pointer events.
2. Click a word. SubMiner selects it using Unicode-aware word boundary detection (`Intl.Segmenter`).
3. Yomitan detects the text selection and opens its popup with dictionary results.
4. From the Yomitan popup, you can add the word directly to Anki.

### On the Invisible Overlay

1. The invisible layer sits over mpv's own subtitle text.
2. Click on any word in the subtitle — SubMiner maps your click position to the underlying text.
3. On macOS, word selection happens automatically on hover.
4. Yomitan popup appears for lookup and card creation.

## Creating Anki Cards

There are three ways to create cards, depending on your workflow.

### 1. Auto-Update from Yomitan

This is the most common flow. Yomitan creates a card in Anki, and SubMiner detects it via polling and enriches it automatically.

1. Click a word → Yomitan popup appears.
2. Click the Anki icon in Yomitan to add the word.
3. SubMiner detects the new card (polls AnkiConnect every 3 seconds by default).
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

This is useful when auto-update polling is disabled or when you want explicit control over which subtitle line gets attached to the card.

| Shortcut                    | Action                                    | Config key                            |
| --------------------------- | ----------------------------------------- | ------------------------------------- |
| `Ctrl/Cmd+C`               | Copy current subtitle                     | `shortcuts.copySubtitle`              |
| `Ctrl/Cmd+Shift+C` + digit | Copy multiple recent lines                | `shortcuts.copySubtitleMultiple`      |
| `Ctrl/Cmd+V`               | Update last card from clipboard           | `shortcuts.updateLastCardFromClipboard` |

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

### What Gets Merged

- **Sentence fields**: Both sentences kept, marked with `[Original]` and `[Duplicate]`.
- **Audio fields**: Both audio clips preserved as separate `[sound:...]` entries.
- **Image fields**: Both images preserved.

Configure in `ankiConnect.isKiku`. See [Anki Integration](/anki-integration#field-grouping-kiku) for the full reference.

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

When enabled, SubMiner highlights words you already know in your Anki deck, making it easier to spot new (N+1) vocabulary during immersion.

### How It Works

1. SubMiner periodically syncs with Anki to build a local cache of known words (expressions/headwords from your configured decks)
2. As subtitles appear, known words are visually highlighted in the visible overlay
3. Unknown words remain unhighlighted — these are your potential mining targets

### Enabling N+1 Mode

```json
{
  "ankiConnect": {
    "nPlusOne": {
      "highlightEnabled": true,
      "refreshMinutes": 1440,
      "matchMode": "headword",
      "minSentenceWords": 3,
      "decks": ["Learning::Japanese"]
    }
  }
}
```

| Option             | Description                                                                         |
| ------------------ | ----------------------------------------------------------------------------------- |
| `highlightEnabled` | Turn on/off the highlighting feature                                                |
| `refreshMinutes`   | How often to refresh the known-word cache (default: 1440 = daily)                   |
| `matchMode`        | `"headword"` (dictionary form) or `"surface"` (exact text match)                    |
| `minSentenceWords` | Minimum sentence length in tokens required to allow N+1 highlighting (default: `3`) |
| `decks`            | Which Anki decks to consider "known" (empty = uses `ankiConnect.deck`)              |

### Use Cases

- **Immersion tracking**: Quickly identify which sentences contain only known words vs. those with new vocabulary
- **Mining focus**: Target sentences with exactly one unknown word (true N+1)
- **Progress visualization**: See your growing vocabulary visually represented in real content

### Immersion Tracking Storage

Immersion data is persisted to SQLite when enabled in `immersionTracking`:

```json
{
  "immersionTracking": {
    "enabled": true,
    "dbPath": ""
  }
}
```

- `dbPath` can be empty (default) to use SubMiner’s app-data `immersion.sqlite`.
- Set an explicit path to move the database (for backups, cloud syncing, or easier inspection).
