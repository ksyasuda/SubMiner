# Immersion Occurrence Tracking Design

**Problem:** `imm_words` and `imm_kanji` only store global aggregates. They cannot answer "where did this word/kanji appear?" at the anime, episode, timestamp, or subtitle-line level.

**Goals:**
- Map normalized words and kanji back to exact subtitle lines.
- Preserve repeated tokens inside one subtitle line.
- Avoid storing token text repeatedly for each repeated token in the same line.
- Keep the change additive and compatible with current top-word/top-kanji stats.

**Non-Goals:**
- Exact token character offsets inside a subtitle line.
- Full stats UI redesign in the same change.
- Replacing existing aggregate tables or existing vocabulary queries.

## Recommended Approach

Add a normalized subtitle-line table plus counted bridge tables from lines to canonical word and kanji rows. Keep `imm_words` and `imm_kanji` as canonical lexeme aggregates, then link them to `imm_subtitle_lines` through one row per unique lexeme per line with `occurrence_count`.

This preserves total frequency within a line without duplicating token text or needing one row per repeated token. Reverse mapping becomes a simple join from canonical lexeme to line row to video/anime context.

## Data Model

### `imm_subtitle_lines`

One row per recorded subtitle line.

Suggested fields:
- `line_id INTEGER PRIMARY KEY AUTOINCREMENT`
- `session_id INTEGER NOT NULL`
- `event_id INTEGER`
- `video_id INTEGER NOT NULL`
- `anime_id INTEGER`
- `line_index INTEGER NOT NULL`
- `segment_start_ms INTEGER`
- `segment_end_ms INTEGER`
- `text TEXT NOT NULL`
- `CREATED_DATE INTEGER`
- `LAST_UPDATE_DATE INTEGER`

Notes:
- `event_id` links back to `imm_session_events` when the subtitle-line event is written.
- `anime_id` is nullable because some rows may predate anime linkage or come from unresolved media.

### `imm_word_line_occurrences`

One row per normalized word per subtitle line.

Suggested fields:
- `line_id INTEGER NOT NULL`
- `word_id INTEGER NOT NULL`
- `occurrence_count INTEGER NOT NULL`
- `PRIMARY KEY(line_id, word_id)`

`word_id` points at the canonical row in `imm_words`.

### `imm_kanji_line_occurrences`

One row per kanji per subtitle line.

Suggested fields:
- `line_id INTEGER NOT NULL`
- `kanji_id INTEGER NOT NULL`
- `occurrence_count INTEGER NOT NULL`
- `PRIMARY KEY(line_id, kanji_id)`

`kanji_id` points at the canonical row in `imm_kanji`.

## Write Path

During `recordSubtitleLine(...)`:

1. Normalize and validate the line as today.
2. Compute counted word and kanji occurrences for the line.
3. Upsert canonical `imm_words` / `imm_kanji` rows as today.
4. Insert one `imm_subtitle_lines` row for the line.
5. Insert counted bridge rows for each normalized word and kanji found in that line.

Counting rules:
- Words: count repeated allowed tokens in the token list; skip tokens excluded by the existing POS/noise filter.
- Kanji: count repeated kanji characters from the visible subtitle line text.

## Query Shape

Add reverse-mapping query functions for:
- word -> recent occurrence rows
- kanji -> recent occurrence rows

Each row should include enough context for drilldown:
- anime id/title
- video id/title
- session id
- line index
- segment start/end
- subtitle text
- occurrence count within that line

Existing top-word/top-kanji aggregate queries stay in place.

## Edge Cases

- Repeated tokens in one line: store once per lexeme per line with `occurrence_count > 1`.
- Duplicate identical lines in one session: each subtitle event gets its own `imm_subtitle_lines` row.
- No anime link yet: keep `anime_id` null and still preserve the line/video/session mapping.
- Legacy DBs: additive migration only; no destructive rebuild of existing word/kanji data.

## Testing Strategy

Start with focused DB-backed tests:
- schema test for new line/bridge tables and indexes
- service test for counted word/kanji line persistence
- query tests for reverse mapping from word/kanji to line/anime/video context
- migration test for existing DBs gaining the new tables cleanly

Primary verification lane: `bun run test:immersion:sqlite:src`, then broader lanes only if API/runtime surfaces widen.
