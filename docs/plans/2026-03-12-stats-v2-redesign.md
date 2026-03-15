# Stats Dashboard v2 Redesign

## Summary

Redesign the stats dashboard to focus on session/media history as the primary experience. Activity feed as the default view, dedicated Library tab with anime cover art (via Anilist API), per-anime drill-down pages, and bug fixes for watch time inflation and relative date formatting.

## Bug Fixes (Pre-requisite)

### Watch Time Inflation

Telemetry values (`active_watched_ms`, `total_watched_ms`, `lines_seen`, `words_seen`, etc.) are cumulative snapshots — each sample stores the running total for that session. Both `getSessionSummaries` (query.ts) and `upsertDailyRollupsForGroups` / `upsertMonthlyRollupsForGroups` (maintenance.ts) incorrectly use `SUM()` across all telemetry rows instead of `MAX()` per session.

**Fix:**
- `getSessionSummaries`: change `SUM(t.active_watched_ms)` → `MAX(t.active_watched_ms)` (already grouped by `s.session_id`)
- `upsertDailyRollupsForGroups` / `upsertMonthlyRollupsForGroups`: use a subquery that gets `MAX()` per session_id, then `SUM()` across sessions
- Run `forceRebuild` rollup after migration to recompute all rollups

### Relative Date Formatting

`formatRelativeDate` only has day-level granularity ("Today", "Yesterday"). Add minute and hour levels:
- < 1 min → "just now"
- < 60 min → "Xm ago"
- < 24 hours → "Xh ago"
- < 2 days → "Yesterday"
- < 7 days → "Xd ago"
- otherwise → locale date string

## Tab Structure

**Overview** (default) | **Library** | **Trends** | **Vocabulary**

### Overview Tab — Activity Feed

Top section: hero stats (watch time today, cards mined today, streak, all-time total hours).

Below: recent sessions listed chronologically, grouped by day headers ("Today", "Yesterday", "March 10"). Each session row shows:
- Small cover art thumbnail (from Anilist cache)
- Clean title with episode info ("The Eminence in Shadow — Episode 5")
- Relative timestamp ("32m ago") and active duration ("24m active")
- Per-session stats: cards mined, words seen

### Library Tab — Cover Art Grid

Grid of anime cover art cards fetched from Anilist API. Each card shows:
- Cover art image (3:4 aspect ratio)
- Episode count badge
- Title, total watch time, cards mined

Controls: search bar, filter chips (All / Watching / Completed), total count + time summary.

Clicking a card navigates to the per-anime detail view.

### Per-Anime Detail View

Navigated from Library card click. Sections:
1. **Header** — cover art, title, total watch time, total episodes, total cards mined, avg session length
2. **Watch time chart** — bar chart scoped to this anime over time (14/30/90d range selector)
3. **Session history** — all sessions for this anime with timestamps, durations, per-session stats
4. **Vocabulary** — words and kanji learned from this show (joined via session events → video_id)

### Trends & Vocabulary Tabs

Keep existing implementation, mostly unchanged for v2.

## Anilist Integration & Cover Art Cache

### Title Parsing

Parse show name from `canonical_title` to search Anilist:
- Jellyfin titles are already clean (use as-is)
- Local file titles: use existing `guessit` + `parseMediaInfo` fallback from `anilist-updater.ts`
- Strip episode info, codec tags, resolution markers via regex

### Anilist GraphQL Query

Search query (no auth needed for public anime search):
```graphql
query ($search: String!) {
  Page(perPage: 5) {
    media(search: $search, type: ANIME) {
      id
      coverImage { large medium }
      title { romaji english native }
      episodes
    }
  }
}
```

### Cover Art Cache

New SQLite table `imm_media_art`:
- `video_id` INTEGER PRIMARY KEY (FK to imm_videos)
- `anilist_id` INTEGER
- `cover_url` TEXT
- `cover_blob` BLOB (cached image binary)
- `title_romaji` TEXT
- `title_english` TEXT
- `episodes_total` INTEGER
- `fetched_at_ms` INTEGER
- `CREATED_DATE` INTEGER
- `LAST_UPDATE_DATE` INTEGER

Serve cached images via: `GET /api/stats/media/:videoId/cover`

Fallback: gray placeholder with first character of title if no Anilist match.

### Rate Limiting Strategy

**Current state:** Anilist limit is 30 req/min (temporarily reduced from 90). Existing post-watch updater uses up to 3 requests per episode (search + entry lookup + save mutation). Retry queue can also fire requests. No centralized rate limiter.

**Centralized rate limiter:**
- Shared sliding-window tracker (array of timestamps) for all Anilist calls
- App-wide cap: 20 req/min (leaving 10 req/min headroom)
- All callers go through the limiter: existing `anilistGraphQl` helper and new cover art fetcher
- Read `X-RateLimit-Remaining` from response headers; if < 5, pause until window resets
- On 429 response, honor `Retry-After` header

**Cover art fetching behavior:**
- Lazy & one-shot: only fetch when a video appears in the stats UI with no cached art
- Once cached in SQLite, never re-fetch (cover art doesn't change)
- On first Library load with N uncached titles, fetch sequentially with ~3s gap between requests
- Show placeholder for unfetched titles, fill in as fetches complete

## New API Endpoints

- `GET /api/stats/media` — all media with aggregated stats (total time, episodes watched, cards, last watched, cover art status)
- `GET /api/stats/media/:videoId` — single media detail: session history, rollups, vocab for that video
- `GET /api/stats/media/:videoId/cover` — cached cover art image (binary response)

## New Database Queries

- `getMediaLibrary(db)` — group sessions by video_id, aggregate stats, join with imm_media_art
- `getMediaDetail(db, videoId)` — sessions + daily rollups + vocab scoped to one video_id
- `getMediaVocabulary(db, videoId)` — words/kanji from sessions belonging to a specific video_id (join imm_session_events with imm_sessions on video_id)

## Data Flow

```
Library tab loads
  → GET /api/stats/media
  → Returns list of videos with aggregated stats + cover art status
  → For videos without cached art:
    → Background: parse title → search Anilist → download cover → cache in SQLite
    → Rate-limited via centralized sliding window (20 req/min cap)
  → UI shows placeholders, fills in as covers arrive

User clicks anime card
  → GET /api/stats/media/:videoId
  → Returns sessions, rollups, vocab for that video
  → Renders detail view with all four sections
```
