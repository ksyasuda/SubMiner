# Episode Detail & Anki Card Link Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add expandable episode detail rows with per-episode sessions/words/cards, store Anki note IDs in card mined events, and add "Open in Anki" button that opens the card browser.

**Architecture:** Extend the recordCardsMined callback chain to pass note IDs alongside count. Store them in the existing payload_json column. Add a stats server endpoint that proxies AnkiConnect's guiBrowse. Frontend makes episode rows expandable with inline detail content.

**Tech Stack:** Hono (backend API), AnkiConnect (guiBrowse), React + Recharts + Tailwind/Catppuccin (frontend), bun test (backend), Vitest (frontend)

---

## Task 1: Extend recordCardsAdded callback to pass note IDs

**Files:**
- Modify: `src/anki-integration/anki-connect-proxy.ts`
- Modify: `src/anki-integration/polling.ts`
- Modify: `src/anki-integration.ts`

**Step 1: Update the callback type**

In `src/anki-integration/anki-connect-proxy.ts` line 18, change:
```typescript
recordCardsAdded?: (count: number) => void;
```
to:
```typescript
recordCardsAdded?: (count: number, noteIds: number[]) => void;
```

In `src/anki-integration/polling.ts` line 12, same change.

**Step 2: Pass note IDs through the proxy callback**

In `src/anki-integration/anki-connect-proxy.ts` line 349, change:
```typescript
this.deps.recordCardsAdded?.(enqueuedCount);
```
to:
```typescript
this.deps.recordCardsAdded?.(enqueuedCount, noteIds.filter(id => !this.pendingNoteIdSet.has(id)));
```

Wait — the dedup already happened by this point. The `noteIds` param to `enqueueNotes` contains the raw IDs. We need the ones that were actually enqueued (not filtered out as duplicates). Track them:

Actually, look at lines 334-348: it iterates `noteIds`, skips duplicates, and pushes accepted ones to `this.pendingNoteIds`. The `enqueuedCount` tracks how many were accepted. We need to collect those IDs:

```typescript
enqueueNotes(noteIds: number[]): void {
  const accepted: number[] = [];
  for (const noteId of noteIds) {
    if (this.pendingNoteIdSet.has(noteId) || this.inFlightNoteIds.has(noteId)) {
      continue;
    }
    this.pendingNoteIds.push(noteId);
    this.pendingNoteIdSet.add(noteId);
    accepted.push(noteId);
  }
  if (accepted.length > 0) {
    this.deps.recordCardsAdded?.(accepted.length, accepted);
  }
  // ... rest of method
}
```

**Step 3: Pass note IDs through the polling callback**

In `src/anki-integration/polling.ts` line 84, change:
```typescript
this.deps.recordCardsAdded?.(newNoteIds.length);
```
to:
```typescript
this.deps.recordCardsAdded?.(newNoteIds.length, newNoteIds);
```

**Step 4: Update AnkiIntegration callback chain**

In `src/anki-integration.ts`:

Line 140, change field type:
```typescript
private recordCardsMinedCallback: ((count: number, noteIds?: number[]) => void) | null = null;
```

Line 154, update constructor param:
```typescript
recordCardsMined?: (count: number, noteIds?: number[]) => void
```

Lines 214-216 (polling deps), change to:
```typescript
recordCardsAdded: (count, noteIds) => {
  this.recordCardsMinedCallback?.(count, noteIds);
}
```

Lines 238-240 (proxy deps), same change.

Line 1125-1127 (setter), update signature:
```typescript
setRecordCardsMinedCallback(callback: ((count: number, noteIds?: number[]) => void) | null): void
```

**Step 5: Commit**

```bash
git commit -m "feat(anki): pass note IDs through recordCardsAdded callback chain"
```

---

## Task 2: Store note IDs in card mined event payload

**Files:**
- Modify: `src/core/services/immersion-tracker-service.ts`

**Step 1: Update recordCardsMined to accept and store note IDs**

Find the `recordCardsMined` method (line 759). Change signature and payload:

```typescript
recordCardsMined(count = 1, noteIds?: number[]): void {
  if (!this.sessionState) return;
  this.sessionState.cardsMined += count;
  this.sessionState.pendingTelemetry = true;
  this.recordWrite({
    kind: 'event',
    sessionId: this.sessionState.sessionId,
    sampleMs: Date.now(),
    eventType: EVENT_CARD_MINED,
    wordsDelta: 0,
    cardsDelta: count,
    payloadJson: sanitizePayload(
      { cardsMined: count, ...(noteIds?.length ? { noteIds } : {}) },
      this.maxPayloadBytes,
    ),
  });
}
```

**Step 2: Update the caller in main.ts**

Find where `recordCardsMined` is called (around line 2506-2508 and 3409-3411). Pass through noteIds:

```typescript
recordCardsMined: (count, noteIds) => {
  ensureImmersionTrackerStarted();
  appState.immersionTracker?.recordCardsMined(count, noteIds);
}
```

**Step 3: Commit**

```bash
git commit -m "feat(immersion): store anki note IDs in card mined event payload"
```

---

## Task 3: Add episode-level query functions

**Files:**
- Modify: `src/core/services/immersion-tracker/query.ts`
- Modify: `src/core/services/immersion-tracker/types.ts`
- Modify: `src/core/services/immersion-tracker-service.ts`

**Step 1: Add types**

In `types.ts`, add:
```typescript
export interface EpisodeCardEventRow {
  eventId: number;
  sessionId: number;
  tsMs: number;
  cardsDelta: number;
  noteIds: number[];
}
```

**Step 2: Add query functions**

In `query.ts`:

```typescript
export function getEpisodeWords(db: DatabaseSync, videoId: number, limit = 50): AnimeWordRow[] {
  return db.prepare(`
    SELECT w.id AS wordId, w.headword, w.word, w.reading, w.part_of_speech AS partOfSpeech,
           SUM(o.occurrence_count) AS frequency
    FROM imm_word_line_occurrences o
    JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
    JOIN imm_words w ON w.id = o.word_id
    WHERE sl.video_id = ?
    GROUP BY w.id
    ORDER BY frequency DESC
    LIMIT ?
  `).all(videoId, limit) as unknown as AnimeWordRow[];
}

export function getEpisodeSessions(db: DatabaseSync, videoId: number): SessionSummaryQueryRow[] {
  return db.prepare(`
    SELECT
      s.session_id AS sessionId, s.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      s.started_at_ms AS startedAtMs, s.ended_at_ms AS endedAtMs,
      COALESCE(MAX(t.total_watched_ms), 0) AS totalWatchedMs,
      COALESCE(MAX(t.active_watched_ms), 0) AS activeWatchedMs,
      COALESCE(MAX(t.lines_seen), 0) AS linesSeen,
      COALESCE(MAX(t.words_seen), 0) AS wordsSeen,
      COALESCE(MAX(t.tokens_seen), 0) AS tokensSeen,
      COALESCE(MAX(t.cards_mined), 0) AS cardsMined,
      COALESCE(MAX(t.lookup_count), 0) AS lookupCount,
      COALESCE(MAX(t.lookup_hits), 0) AS lookupHits
    FROM imm_sessions s
    JOIN imm_videos v ON v.video_id = s.video_id
    LEFT JOIN imm_session_telemetry t ON t.session_id = s.session_id
    WHERE s.video_id = ?
    GROUP BY s.session_id
    ORDER BY s.started_at_ms DESC
  `).all(videoId) as SessionSummaryQueryRow[];
}

export function getEpisodeCardEvents(db: DatabaseSync, videoId: number): EpisodeCardEventRow[] {
  const rows = db.prepare(`
    SELECT e.event_id AS eventId, e.session_id AS sessionId,
           e.ts_ms AS tsMs, e.cards_delta AS cardsDelta,
           e.payload_json AS payloadJson
    FROM imm_session_events e
    JOIN imm_sessions s ON s.session_id = e.session_id
    WHERE s.video_id = ? AND e.event_type = 4
    ORDER BY e.ts_ms DESC
  `).all(videoId) as Array<{ eventId: number; sessionId: number; tsMs: number; cardsDelta: number; payloadJson: string | null }>;

  return rows.map(row => {
    let noteIds: number[] = [];
    if (row.payloadJson) {
      try {
        const parsed = JSON.parse(row.payloadJson);
        if (Array.isArray(parsed.noteIds)) noteIds = parsed.noteIds;
      } catch {}
    }
    return { eventId: row.eventId, sessionId: row.sessionId, tsMs: row.tsMs, cardsDelta: row.cardsDelta, noteIds };
  });
}
```

**Step 3: Add wrapper methods to immersion-tracker-service.ts**

**Step 4: Commit**

```bash
git commit -m "feat(stats): add episode-level query functions for sessions, words, cards"
```

---

## Task 4: Add episode detail and Anki browse API endpoints

**Files:**
- Modify: `src/core/services/stats-server.ts`
- Modify: `src/core/services/__tests__/stats-server.test.ts`

**Step 1: Add episode detail endpoint**

```typescript
app.get('/api/stats/episode/:videoId/detail', async (c) => {
  const videoId = parseIntQuery(c.req.param('videoId'), 0);
  if (videoId <= 0) return c.body(null, 400);
  const sessions = await tracker.getEpisodeSessions(videoId);
  const words = await tracker.getEpisodeWords(videoId);
  const cardEvents = await tracker.getEpisodeCardEvents(videoId);
  return c.json({ sessions, words, cardEvents });
});
```

**Step 2: Add Anki browse endpoint**

```typescript
app.post('/api/stats/anki/browse', async (c) => {
  const noteId = parseIntQuery(c.req.query('noteId'), 0);
  if (noteId <= 0) return c.body(null, 400);
  try {
    const response = await fetch('http://127.0.0.1:8765', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'guiBrowse', version: 6, params: { query: `nid:${noteId}` } }),
    });
    const result = await response.json();
    return c.json(result);
  } catch (err) {
    return c.json({ error: 'Failed to reach AnkiConnect' }, 502);
  }
});
```

**Step 3: Add tests and verify**

Run: `bun test ./src/core/services/__tests__/stats-server.test.ts`

**Step 4: Commit**

```bash
git commit -m "feat(stats): add episode detail and anki browse endpoints"
```

---

## Task 5: Add frontend types and API client methods

**Files:**
- Modify: `stats/src/types/stats.ts`
- Modify: `stats/src/lib/api-client.ts`
- Modify: `stats/src/lib/ipc-client.ts`

**Step 1: Add types**

```typescript
export interface EpisodeCardEvent {
  eventId: number;
  sessionId: number;
  tsMs: number;
  cardsDelta: number;
  noteIds: number[];
}

export interface EpisodeDetailData {
  sessions: SessionSummary[];
  words: AnimeWord[];
  cardEvents: EpisodeCardEvent[];
}
```

**Step 2: Add API client methods**

```typescript
getEpisodeDetail: (videoId: number) => fetchJson<EpisodeDetailData>(`/api/stats/episode/${videoId}/detail`),
ankiBrowse: (noteId: number) => fetchJson<unknown>(`/api/stats/anki/browse?noteId=${noteId}`, { method: 'POST' }),
```

Mirror in ipc-client.

**Step 3: Commit**

```bash
git commit -m "feat(stats): add episode detail types and API client methods"
```

---

## Task 6: Build EpisodeDetail component

**Files:**
- Create: `stats/src/components/anime/EpisodeDetail.tsx`
- Modify: `stats/src/components/anime/EpisodeList.tsx`
- Modify: `stats/src/components/anime/AnimeCardsList.tsx`

**Step 1: Create EpisodeDetail component**

Inline expandable content showing:
- Sessions list (compact: time, duration, cards, words)
- Cards mined list with "Open in Anki" button per note ID
- Top words grid (reuse AnimeWordList pattern)

Fetches data from `getEpisodeDetail(videoId)` on mount.

"Open in Anki" button calls `apiClient.ankiBrowse(noteId)`.

**Step 2: Make EpisodeList rows expandable**

Add `expandedVideoId` state. Clicking a row toggles expansion. Render `EpisodeDetail` below the expanded row.

**Step 3: Make AnimeCardsList rows expandable**

Same pattern — clicking an episode row expands to show `EpisodeDetail`.

**Step 4: Commit**

```bash
git commit -m "feat(stats): add expandable episode detail with anki card links"
```

---

## Task 7: Build and verify

**Step 1: Type check**
Run: `npx tsc --noEmit`

**Step 2: Run backend tests**
Run: `bun test ./src/core/services/__tests__/stats-server.test.ts`

**Step 3: Run frontend tests**
Run: `npx vitest run`

**Step 4: Build**
Run: `npx vite build`

**Step 5: Commit any fixes**

```bash
git commit -m "feat(stats): episode detail and anki link complete"
```
