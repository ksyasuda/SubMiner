# Immersion Anime Metadata Design

**Problem:** The immersion database is keyed around videos and sessions, which makes it awkward to present anime-centric stats such as per-anime totals, episode progress, and season breakdowns. We need first-class anime metadata without requiring migration or backfill support for existing databases.

**Goals:**
- Add anime-level identity that can be shared across multiple video files and rewatches.
- Persist parsed episode/season metadata so stats can group by anime, season, and episode.
- Use existing filename parsing conventions: `guessit` first, built-in parser fallback.
- Create provisional anime rows even when AniList lookup fails.
- Keep the change additive and forward-looking; do not spend time on migrations/backfill.

**Non-Goals:**
- Backfilling or migrating existing user databases.
- Perfect anime identity resolution across every edge case.
- Building the entire new stats UI in this design doc.
- Replacing existing `canonical_title` or current video/session APIs immediately.

## Recommended Approach

Add a new `imm_anime` table for anime-level metadata and link each `imm_videos` row to one anime row through `anime_id`. Keep season/episode and filename-derived fields on `imm_videos`, because those belong to a concrete file, not the anime as a whole.

Anime rows should exist even when AniList lookup fails. In that case, use a normalized parsed-title key as provisional identity. If the same anime is resolved to AniList later, upgrade the existing anime row in place instead of creating a duplicate.

## Data Model

### `imm_anime`

One row per anime identity.

Suggested fields:
- `anime_id INTEGER PRIMARY KEY AUTOINCREMENT`
- `identity_key TEXT NOT NULL UNIQUE`
- `parsed_title TEXT NOT NULL`
- `normalized_title TEXT NOT NULL`
- `anilist_id INTEGER`
- `title_romaji TEXT`
- `title_english TEXT`
- `title_native TEXT`
- `episodes_total INTEGER`
- `parser_source TEXT`
- `parser_confidence TEXT`
- `metadata_json TEXT`
- `CREATED_DATE INTEGER`
- `LAST_UPDATE_DATE INTEGER`

Identity rules:
- Resolved anime: `identity_key = anilist:<id>`
- Provisional anime: `identity_key = title:<normalized parsed title>`
- When a provisional row later gets an AniList match, update that row's `identity_key` to `anilist:<id>` and fill AniList metadata.

### `imm_videos`

Keep existing video metadata. Add:
- `anime_id INTEGER`
- `parsed_filename TEXT`
- `parsed_title TEXT`
- `parsed_title_normalized TEXT`
- `parsed_season INTEGER`
- `parsed_episode INTEGER`
- `parsed_episode_title TEXT`
- `parser_source TEXT`
- `parser_confidence TEXT`
- `parse_metadata_json TEXT`

`canonical_title` remains for compatibility. New fields are additive.

## Parsing and Lookup Flow

During `handleMediaChange(...)`:

1. Normalize path/title with the existing tracker flow.
2. Build/create the video row as today.
3. Parse anime metadata:
   - use `guessit` against the basename/title when available
   - fallback to existing `parseMediaInfo`
4. Use the parsed title to create/find a provisional anime row if needed.
5. Attempt AniList lookup using the same guessit-first, fallback-parser approach already used elsewhere.
6. If AniList lookup succeeds:
   - upgrade or fill the anime row with AniList id/title metadata
   - keep per-video season/episode fields on the video row
7. Link the video row to `anime_id` and store parsed per-video metadata.

## Query Shape

Add anime-aware query functions without deleting current video/session queries:
- anime library list
- anime detail summary
- anime episode list / season breakdown
- anime sessions list

Aggregation should group by `anime_id`, not `canonical_title`, so rewatches and multiple files collapse correctly.

## Edge Cases

- Multiple files for one anime: many videos may point to one anime row.
- Rewatches: same video/session history still aggregates under one anime row.
- No AniList match: keep provisional anime row keyed by normalized parsed title.
- Later AniList match: upgrade provisional row in place.
- Parser disagreement between files: season/episode remain per-video; anime identity uses AniList id or normalized parsed title.
- Remote/Jellyfin playback: use the effective title/path available to the current tracker flow and run the same parser pipeline.

## Testing Strategy

Start red/green with focused DB-backed tests:
- schema test for `imm_anime` and new video columns
- storage test for provisional anime creation, reuse, and AniList upgrade
- service test for media-change ingest wiring
- query test for anime-level aggregation and episode breakdown

Primary verification lane for implementation: `bun run test:immersion:sqlite:src`, then broader repo verification as needed.
