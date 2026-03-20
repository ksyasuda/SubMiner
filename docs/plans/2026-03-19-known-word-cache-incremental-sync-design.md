<!-- read_when: changing known-word cache lifecycle, stats cache semantics, or Anki sync behavior -->

# Incremental Known-Word Cache Sync

## Goal

Stop rebuilding the entire known-word cache on startup or routine refreshes. Keep the cache correct through incremental reconciliation on the configured sync cadence, with an immediate append path for freshly mined cards.

## Scope

- Persist per-note extracted known-word snapshots beside the existing global `words` list.
- Replace startup refresh with load-only behavior.
- Make timed refresh diff current Anki note IDs against cached note IDs, then apply add/remove/edit deltas.
- Add `ankiConnect.knownWords.addMinedWordsImmediately`, default `true`.
- Keep full rebuild out of normal lifecycle; reserve it for explicit doctor tooling.

## Data Model

Persist versioned cache state with:

- `words`: deduplicated global known-word set for stats/UI consumers
- `notes`: record of `noteId -> extractedWords[]`
- `refreshedAtMs`
- `scope`

The in-memory manager derives the global set from the per-note snapshots during sync updates so deletes and edits can remove stale words safely.

## Sync Behavior

- Startup: load persisted state only
- Interval tick or explicit refresh command: run incremental sync
- Incremental sync:
  - query tracked note IDs for configured deck scope
  - remove note snapshots for note IDs that disappeared
  - fetch `notesInfo` for note IDs that are new or need field reconciliation
  - compare extracted words per note and update the global set

## Immediate Mining Path

When SubMiner already has fresh `noteInfo` after mining or updating a note, append/update that note snapshot immediately if `addMinedWordsImmediately` is enabled.

## Verification

- focused cache manager tests for add/delete/edit reconciliation
- focused integration/config tests for startup behavior and new config flag
- config verification lane because defaults/schema/example change
