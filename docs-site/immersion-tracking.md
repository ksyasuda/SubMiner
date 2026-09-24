# Immersion tracking

SubMiner records what you watch and mine in a local SQLite database and shows it in a stats dashboard. Tracking is on by default. Nothing leaves your machine.

## What gets tracked

- Watch sessions: time watched, subtitle lines seen, words seen, cards mined, pauses and seeks.
- Every primary subtitle line you see, with its timing, so you can search and mine from it later.
- Vocabulary and kanji you encounter, with how often and where.
- Library entries per show and episode, with cover art from AniList (or TMDB for live action) and YouTube channel metadata.

An episode counts as watched once you reach 85% of it.

## Setup

Tracking needs no setup. To turn it off or move the database:

```jsonc
{
  "immersionTracking": {
    "enabled": true,
    "dbPath": "",
  },
}
```

An empty `dbPath` stores `immersion.sqlite` in SubMiner's config directory (`~/.config/SubMiner/` on Linux). Set a path to keep it elsewhere.

To share stats and watch history between machines, use [`subminer sync <host>`](/launcher-script#sync-between-machines). It merges both databases. Copying the file with a cloud sync tool makes one side overwrite the other.

## Open the dashboard

- In the overlay: focus it and press the `stats.toggleKey` key (Backquote by default).
- In a browser: run `subminer stats`, then open `http://127.0.0.1:6969` (or your `stats.serverPort`). Set `stats.autoOpenBrowser` to open it automatically.
- Background server: `subminer stats -b` starts a stats server that keeps running without the launcher attached. `subminer stats -s` stops it. You can still start SubMiner for playback while it runs.

`subminer stats` fails if `immersionTracking.enabled` is `false`. The server only answers on localhost, so reverse proxies and Tailscale Serve URLs do not work.

## Stats dashboard

### Overview

Recent sessions, a streak calendar, watch-time history, and totals for completed episodes and shows.

![Stats Overview](/screenshots/stats-overview.png)

### Library

Your shows as cover-art cards with search, sorting, per-series progress, and an episode list linking to mined cards. The **All Titles** / **Anime** / **Live Action** / **YouTube** selector filters the grid. YouTube videos are grouped by channel.

Seasons get separate cards when a season number is detected. Live-action titles that AniList cannot match are looked up on [TMDB](/configuration#tmdb). If a title gets no match, open it and use **Link to TMDB** to pick one by hand.

The same show can end up on several cards when release names disagree. To fix that:

- Merge: click **Select**, tick the duplicate cards, choose **Merge Selected**, and pick the entry to keep. Sessions, cards, and watch time move over, and future episodes with those names join the kept entry.
- Move one episode: hover its row in the episode list and click **→** to assign it to another entry. SubMiner remembers the correction.
- Suggested merges appear as **Possible duplicate** above the grid. Choose **Review merge** or **Not duplicates**.

**Delete Entry** in a title's header removes the show with all its episodes, sessions, and lines. You cannot delete the title that is currently playing.

![Stats Library](/screenshots/stats-library.png)

### Trends

Charts for watch time, cards, words, and sessions per day or month, running totals, efficiency (words per minute, cards per hour), and viewing patterns by weekday and hour. Each chart has its own date range and grouping.

![Stats Trends](/screenshots/stats-trends.png)

### Sessions

Session history with new-word activity and pause, seek, and card markers. The **↗** button on a row opens that show's detail view.

![Stats Sessions](/screenshots/stats-sessions.png)

### Vocabulary

Unique words and kanji you have seen, new words per day, frequency rank tables with Hide Known and Hide Kana filters, and a kanji breakdown. Click a word to see every line it appeared in.

- **Exclusions** hides words from every vocabulary view. You can restore them from the same dialog.
- **Duplicates** cleans up lines repeated by karaoke openings and animated signs (see [Repeated lines](#repeated-lines)).

![Stats Vocabulary](/screenshots/stats-vocabulary.png)

### Search

Searches the primary subtitle lines and titles in your history. **Search by headword** is on by default, so `知らない` also finds inflected forms. Turn it off for exact text matching. Secondary subtitles are not searched.

## Mining from the dashboard

Search results and the Vocabulary word panel can create cards from past lines, as long as the source video file is still available:

- **Mine Word**: full Yomitan lookup for the word, plus sentence, audio, and image.
- **Mine Sentence**: a sentence card with `IsSentenceCard` set, for Lapis and Kiku note types.
- **Mine Audio**: an audio card with `IsAudioCard` set.

Word and audio mining appear only when the word occurs in the sentence. All three use your `ankiConnect` deck, note type, fields, and media settings. Anki must be running, and Mine Word needs Yomitan dictionaries.

## Repeated lines

Karaoke openings and animated signs repeat the same text once per frame. SubMiner collapses these as it records, so one lyric is stored once. Stats recorded before that can hold hundreds of copies and skew Top Repeated Words.

To clean them, use **Duplicates** in the Vocabulary tab: pick a time window, **Scan** to preview, then **Clean Up**. Or from the terminal:

```bash
subminer stats cleanup --duplicate-lines --dry-run --lookback-days 30
subminer stats cleanup --duplicate-lines --lookback-days 30
```

Leave out `--lookback-days` to scan all history. Word and kanji counts are corrected. Watch time and session totals are not changed.

## Maintenance commands

| Command                                    | What it does                                                              |
| ------------------------------------------ | ------------------------------------------------------------------------- |
| `subminer stats cleanup`                   | Repair word readings and part of speech, drop words that fail the filters |
| `subminer stats cleanup -l`                | Recompute lifetime totals from episode history, keeping old totals        |
| `subminer stats cleanup --duplicate-lines` | Collapse repeated karaoke and sign lines (see above)                      |

`subminer stats rebuild` and `subminer stats backfill` run the same lifetime repair as `cleanup -l`.

## Retention

By default SubMiner keeps everything. To limit history, set `immersionTracking.retentionPreset` to `minimal`, `balanced`, or `deep-history`, or set the `immersionTracking.retention.*Days` values yourself (`0` keeps all). Lifetime totals and vocabulary counts are stored separately and stay exact when old sessions are pruned.

See [Immersion tracking](/configuration#immersion-tracking) and [Stats dashboard](/configuration#stats-dashboard) in the config reference for every option and default.
