type: fixed
area: sync

- Fixed word/kanji frequencies double-counting across syncs when the remote snapshot contained a stale active session (e.g. after a crash): a word new to the local machine adopted the remote's full lifetime frequency, which already included the active session's partial occurrences, and those occurrences were added again when the session finalized and synced. Newly adopted words/kanji now exclude active-session counts, which arrive once the session completes.
