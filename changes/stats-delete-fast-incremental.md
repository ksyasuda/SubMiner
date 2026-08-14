type: fixed
area: stats

- Stats deletes no longer freeze the stats dashboard: the delete worker module now resolves when running from source, so deletes actually run off the serving thread instead of silently falling back to it.
- Deletes now subtract their exact contribution from lifetime summaries instead of rebuilding them from retained sessions, making delete cost proportional to what is deleted and preserving lifetime totals older than the session retention window.
- If the delete worker crashes, the delete now retries on the current thread instead of failing.
- Library merges, video moves, AniList reassignments, and `subminer stats cleanup -l` also stopped rebuilding lifetime summaries from retained sessions; they now recompute from per-episode history, so those operations are faster and no longer erase lifetime totals older than the session retention window.
- Deleting content that contains very common words no longer rescans every occurrence of those words across the whole library; first/last-seen dates are refreshed with index seeks instead.
