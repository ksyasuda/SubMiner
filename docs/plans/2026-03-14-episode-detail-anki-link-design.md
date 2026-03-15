# Episode Detail & Anki Card Link — Design

**Date**: 2026-03-14
**Status**: Approved

## Motivation

The anime detail page shows episodes and cards mined but lacks drill-down into individual episodes. Users want to see per-episode stats (sessions, words, cards) and link directly to mined Anki cards.

## Design

### 1. Episode Expandable Detail

Click an episode row in `EpisodeList` or `AnimeCardsList` → expands inline:
- Sessions list for this episode (sessions linked to video_id)
- Cards mined list — timestamps + "Open in Anki" button per card (when note ID available)
- Top words from this episode (word occurrences scoped to video_id)

### 2. Anki Note ID Storage

- Extend `recordCardsMined` callback to accept note IDs: `recordCardsMined(count, noteIds)`
- Store in CARD_MINED event payload: `{ cardsMined: 1, noteIds: [12345] }`
- Proxy already has note IDs in `pendingNoteIds` — pass through callback chain
- Polling has note IDs from `newNoteIds` — same treatment
- No schema change — note IDs stored in existing `payload_json` column on `imm_session_events`

### 3. "Open in Anki" Flow

- New endpoint: `POST /api/stats/anki/browse?noteId=12345`
- Calls AnkiConnect `guiBrowse` with query `nid:12345`
- Opens Anki's card browser filtered to that note
- Frontend button hits this endpoint

### 4. Episode Words

- New query: `getEpisodeWords(videoId)` — like `getAnimeWords` but filtered by video_id
- Reuse AnimeWordList component pattern

### 5. Backend Changes

**Modified files:**
- `src/anki-integration/anki-connect-proxy.ts` — pass note IDs through recordCardsAdded callback
- `src/anki-integration/polling.ts` — pass note IDs through recordCardsAdded callback
- `src/anki-integration.ts` — update callback signature
- `src/core/services/immersion-tracker-service.ts` — accept and store note IDs in recordCardsMined
- `src/core/services/immersion-tracker/query.ts` — add getEpisodeWords, getEpisodeSessions, getEpisodeCardEvents
- `src/core/services/stats-server.ts` — add episode detail and anki browse endpoints

### 6. Frontend Changes

**Modified files:**
- `stats/src/components/anime/EpisodeList.tsx` — make rows expandable
- `stats/src/components/anime/AnimeCardsList.tsx` — make rows expandable

**New files:**
- `stats/src/components/anime/EpisodeDetail.tsx` — inline expandable content
